#!/usr/bin/env python3
"""
Web UI for pdf2docx: upload a PDF, convert it on the CPU, download the .docx.

The model is loaded once at startup and jobs are processed one at a time.
Start it with docker/start-web.ps1 (see README.md), then open http://localhost:8080
"""
from __future__ import annotations

import base64
import json
import os
import queue
import re
import secrets
import shutil
import threading
import time
import uuid
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from pathlib import Path

import pymupdf
import uvicorn
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, Response

import pdf2docx as core

ROOT = Path(__file__).resolve().parent
JOBS_DIR = Path(os.environ.get("JOBS_DIR", ROOT / "webdata"))
MODEL_DIR = Path(os.environ.get("MODEL_DIR", ROOT / "model"))
MAX_UPLOAD = int(os.environ.get("MAX_UPLOAD_MB", "200")) * 1024 * 1024
THREADS = int(os.environ["OCR_THREADS"]) if os.environ.get("OCR_THREADS") else None
ROTATE = os.environ.get("OCR_ROTATE", "1") != "0"   # re-read sideways / plan pages turned 90 degrees
# Set OCR_BASE_URL to a running vLLM server (e.g. http://host.docker.internal:8000/v1) to
# do the OCR on its GPU instead of this container's CPU: seconds per page instead of minutes.
BASE_URL = os.environ.get("OCR_BASE_URL", "").strip()
WORKERS = max(1, int(os.environ.get("OCR_WORKERS", "6"))) if BASE_URL else 1

WEB_USER = os.environ.get("WEB_USER", "admin")
WEB_PASSWORD = os.environ.get("WEB_PASSWORD", "")          # empty = no login required

app = FastAPI(title="PDF to Word (Thai OCR)", docs_url=None, redoc_url=None, openapi_url=None)


@app.middleware("http")
async def basic_auth(request: Request, call_next):
    """Require HTTP Basic login when WEB_PASSWORD is set (use it whenever the
    site is reachable from other machines)."""
    if WEB_PASSWORD and request.url.path != "/healthz":
        ok = False
        header = request.headers.get("authorization", "")
        if header.lower().startswith("basic "):
            try:
                user, _, password = base64.b64decode(header[6:]).decode("utf-8").partition(":")
                ok = (secrets.compare_digest(user.encode(), WEB_USER.encode())
                      and secrets.compare_digest(password.encode(), WEB_PASSWORD.encode()))
            except (ValueError, UnicodeDecodeError):
                ok = False
        if not ok:
            return Response("Login required", status_code=401,
                            headers={"WWW-Authenticate": 'Basic realm="pdf2docx", charset="UTF-8"'})
    return await call_next(request)


@app.get("/healthz")
def healthz() -> dict:
    return {"model": state["model"]}


jobs: dict[str, dict] = {}
lock = threading.Lock()
todo: queue.Queue[str] = queue.Queue()
state = {"model": "loading", "error": None}       # loading | ready | error

# ---------------------------------------------------------------- job storage

def job_dir(job_id: str) -> Path:
    return JOBS_DIR / job_id


def save(job: dict) -> None:
    tmp = job_dir(job["id"]) / "meta.json.tmp"
    tmp.write_text(json.dumps(job, ensure_ascii=False), encoding="utf-8")
    tmp.replace(job_dir(job["id"]) / "meta.json")


def update(job_id: str, **changes) -> dict | None:
    with lock:
        job = jobs.get(job_id)
        if job is None:                           # deleted while running
            return None
        job.update(changes)
        save(job)
        return dict(job)


def load_existing() -> None:
    """Bring back jobs from a previous run; unfinished ones go back in the queue."""
    JOBS_DIR.mkdir(parents=True, exist_ok=True)
    found = []
    for meta in JOBS_DIR.glob("*/meta.json"):
        try:
            found.append(json.loads(meta.read_text(encoding="utf-8")))
        except (OSError, ValueError):
            continue
    for job in sorted(found, key=lambda j: j.get("created", 0)):
        if job["status"] in ("queued", "running"):
            job["status"] = "queued"
            todo.put(job["id"])
        jobs[job["id"]] = job

# ---------------------------------------------------------------- conversion

def convert(job_id: str, backend) -> None:
    d = job_dir(job_id)
    cache = d / "pages"
    work = d / "tmp"
    cache.mkdir(exist_ok=True)
    work.mkdir(exist_ok=True)
    job = update(job_id, status="running", message="")
    if job is None:
        return
    doc = pymupdf.open(d / "input.pdf")
    try:
        indices = job["indices"]
        results: dict[int, str] = {}
        failed: list[int] = []
        todo_pages = []
        for i in indices:
            cached = cache / f"page_{i + 1:05d}.md"
            text = cached.read_text(encoding="utf-8") if cached.exists() else None
            if text is None or core.FAILED_RE.search(text):
                todo_pages.append(i)
            else:
                results[i] = text
        update(job_id, pages_done=len(results))
        timed = 0
        t0 = time.time()
        pending: dict = {}
        queue_iter = iter(todo_pages)
        with ThreadPoolExecutor(max_workers=WORKERS) as pool:
            while True:
                # render in this thread (PyMuPDF documents are not thread-safe), OCR in the pool
                while len(pending) < WORKERS * 2 and job_id in jobs:
                    i = next(queue_iter, None)
                    if i is None:
                        break
                    img = core.render_page(doc, i, work)
                    pending[pool.submit(core.ocr_page, img, backend, ROTATE)] = (i, img)
                if not pending:
                    break
                finished, _ = wait(pending, return_when=FIRST_COMPLETED)
                for fut in finished:
                    i, img = pending.pop(fut)
                    try:
                        text, ok = fut.result()
                        text = core.fix_thai(text)
                    except Exception as e:        # keep going on single-page failure
                        text, ok = f"> [หน้า {i + 1} ผิดพลาด: {e}]", False
                    img.unlink(missing_ok=True)
                    if ok:
                        (cache / f"page_{i + 1:05d}.md").write_text(text, encoding="utf-8")
                    else:
                        failed.append(i + 1)
                    results[i] = text
                    timed += 1
                    update(job_id, pages_done=len(results), failed=sorted(failed),
                           sec_per_page=round((time.time() - t0) / timed, 1))
        if job_id not in jobs:                    # deleted by the user
            return
        pages_md = [results[i] for i in indices]
        if job_id not in jobs:
            return
        update(job_id, message="กำลังสร้างไฟล์ Word")
        core.markdown_pages_to_docx(pages_md, d / "output.docx", work)
        update(job_id, status="done", message="", finished=time.time())
    finally:
        doc.close()
        shutil.rmtree(work, ignore_errors=True)


def worker() -> None:
    try:
        if BASE_URL:
            backend = core.VllmBackend(BASE_URL, os.environ.get("OCR_MODEL", "nemotron-parse"), 0.8)
            print(f"OCR on the vLLM server at {BASE_URL}, {WORKERS} pages at a time")
        else:
            backend = core.load_cpu_backend(MODEL_DIR, THREADS)
        state["model"] = "ready"
    except Exception as e:
        state.update(model="error", error=f"{type(e).__name__}: {e}")
        return
    while True:
        job_id = todo.get()
        if job_id not in jobs:
            continue
        try:
            convert(job_id, backend)
        except Exception as e:
            update(job_id, status="error", message=f"{type(e).__name__}: {e}")
        if job_id not in jobs:                    # deleted mid-run: remove what was left open
            shutil.rmtree(job_dir(job_id), ignore_errors=True)

# ---------------------------------------------------------------- API

def public(job: dict) -> dict:
    return {k: job.get(k) for k in ("id", "name", "created", "status", "message", "pages_total",
                                    "pages_done", "failed", "sec_per_page", "size")}


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return (ROOT / "web" / "index.html").read_text(encoding="utf-8")


@app.get("/api/jobs")
def list_jobs() -> dict:
    with lock:
        items = sorted(jobs.values(), key=lambda j: j["created"], reverse=True)
        return {"model": state["model"], "model_error": state["error"], "gpu": bool(BASE_URL),
                "max_upload_mb": MAX_UPLOAD // (1024 * 1024),
                "jobs": [public(j) for j in items]}


@app.post("/api/jobs")
def create_job(file: UploadFile = File(...), pages: str = Form("")) -> dict:
    name = Path(file.filename or "document.pdf").name
    if not name.lower().endswith(".pdf"):
        raise HTTPException(400, "รองรับเฉพาะไฟล์ PDF")
    job_id = uuid.uuid4().hex
    d = job_dir(job_id)
    d.mkdir(parents=True)
    try:
        size = 0
        with open(d / "input.pdf", "wb") as out:
            while chunk := file.file.read(1024 * 1024):
                size += len(chunk)
                if size > MAX_UPLOAD:
                    raise HTTPException(413, f"ไฟล์ใหญ่เกิน {MAX_UPLOAD // (1024 * 1024)} MB")
                out.write(chunk)
        try:
            doc = pymupdf.open(d / "input.pdf")
        except Exception:
            raise HTTPException(400, "เปิดไฟล์ PDF ไม่ได้ ไฟล์อาจเสียหาย")
        with doc:
            if not doc.is_pdf:
                raise HTTPException(400, "ไฟล์นี้ไม่ใช่ PDF")
            if doc.needs_pass:
                raise HTTPException(400, "PDF ติดรหัสผ่าน กรุณาปลดรหัสก่อน")
            pages = pages.strip()
            if pages and not re.fullmatch(r"\s*\d+(\s*-\s*\d+)?(\s*,\s*\d+(\s*-\s*\d+)?)*\s*", pages):
                raise HTTPException(400, "รูปแบบช่วงหน้าไม่ถูกต้อง ตัวอย่าง: 1-10,15")
            indices = sorted(set(core.parse_range(pages.replace(" ", "") or None, doc.page_count)))
            indices = [i for i in indices if 0 <= i < doc.page_count]
            if not indices:
                raise HTTPException(400, f"ไม่มีหน้าที่เลือกอยู่ในไฟล์ (ไฟล์มี {doc.page_count} หน้า)")
    except HTTPException:
        shutil.rmtree(d, ignore_errors=True)
        raise
    job = {"id": job_id, "name": name, "created": time.time(), "status": "queued", "message": "",
           "size": size, "indices": indices, "pages_total": len(indices), "pages_done": 0,
           "failed": [], "sec_per_page": None}
    with lock:
        jobs[job_id] = job
        save(job)
    todo.put(job_id)
    return public(job)


def get_job(job_id: str) -> dict:
    if not re.fullmatch(r"[0-9a-f]{32}", job_id) or job_id not in jobs:
        raise HTTPException(404, "ไม่พบงานนี้")
    return jobs[job_id]


@app.get("/api/jobs/{job_id}/download")
def download(job_id: str) -> FileResponse:
    job = get_job(job_id)
    path = job_dir(job_id) / "output.docx"
    if job["status"] != "done" or not path.exists():
        raise HTTPException(409, "ไฟล์ยังแปลงไม่เสร็จ")
    return FileResponse(
        path, filename=Path(job["name"]).stem + ".docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.delete("/api/jobs/{job_id}")
def delete(job_id: str) -> dict:
    get_job(job_id)
    with lock:
        jobs.pop(job_id, None)                    # a running job notices and stops after its page
    shutil.rmtree(job_dir(job_id), ignore_errors=True)
    return {"ok": True}


if __name__ == "__main__":
    load_existing()
    threading.Thread(target=worker, daemon=True).start()
    uvicorn.run(app, host=os.environ.get("HOST", "0.0.0.0"), port=int(os.environ.get("PORT", "8080")),
                log_level="warning")
