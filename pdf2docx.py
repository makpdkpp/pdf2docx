#!/usr/bin/env python3
"""
PDF -> OpenThai2.1-OCR-Markdown (via vLLM) -> MS Word (.docx)

Usage:
    python pdf2docx.py input.pdf -o output.docx
    python pdf2docx.py input.pdf -o output.docx --pages 1-20 --workers 4

Normally run inside the pdf2docx-client Docker image via docker/convert.ps1
(see README.md). To run it directly you need:
    - vLLM server already running
    - pip install pymupdf python-docx markdown, plus the nemotron_parse package
      from the model folder: pip install "./model[client]"
    - pandoc >= 3.0 on PATH
"""
from __future__ import annotations

import argparse
import math
import os
import re
import subprocess
import sys
import tempfile
import time
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from html.parser import HTMLParser
from pathlib import Path

import markdown as md_lib
import pymupdf  # PyMuPDF

MODEL_W, MODEL_H = 1664, 2048          # model input size (width x height)
THAI_FONT = "TH SarabunPSK"                      # the standard face for Thai official documents
BODY_PT = 16

# ---------------------------------------------------------------- PDF -> image

def _scan_zoom(page: pymupdf.Page) -> float | None:
    """If the page is a scan (one image covering most of it), return the zoom
    that renders that image at its own pixel size; otherwise None."""
    area = abs(page.rect)
    best = None
    for info in page.get_image_info():
        box = pymupdf.Rect(info["bbox"]) & page.rect
        if box.is_empty or abs(box) < 0.8 * area or not info.get("width"):
            continue
        zoom = info["width"] / pymupdf.Rect(info["bbox"]).width
        best = max(best or 0, zoom)
    return best


def render_page(doc: pymupdf.Document, index: int, out_dir: Path) -> Path:
    """Render one page so it fits inside 1664x2048 px, keeping aspect ratio.

    Scanned pages are rendered at the scan's own resolution instead: PyMuPDF's
    resampling of an embedded image blurs small text enough to make the model
    misplace table cells. They are never enlarged, and only shrunk (with a
    high-quality filter) when the scan is bigger than the model input.
    """
    page = doc[index]
    w, h = page.rect.width, page.rect.height        # in points
    path = out_dir / f"page_{index + 1:05d}.png"
    zoom = min(MODEL_W / w, MODEL_H / h)
    native = _scan_zoom(page)
    if native is None:
        page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False).save(path)
        return path

    from PIL import Image
    native = min(native, 4 * zoom)                  # guard against huge scans
    pix = page.get_pixmap(matrix=pymupdf.Matrix(native, native), alpha=False)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    if img.width > MODEL_W or img.height > MODEL_H:
        img.thumbnail((MODEL_W, MODEL_H), Image.LANCZOS)
    img.save(path)
    return path


def split_halves(img_path: Path) -> list[Path]:
    """Fallback when a page hits the token limit: split top/bottom with overlap."""
    pix = pymupdf.Pixmap(str(img_path))
    w, h = pix.width, pix.height
    overlap = int(h * 0.04)
    parts = []
    src = pymupdf.open()
    page = src.new_page(width=w, height=h)
    page.insert_image(page.rect, filename=str(img_path))
    for i, (y0, y1) in enumerate([(0, h // 2 + overlap), (h // 2 - overlap, h)]):
        clip = pymupdf.Rect(0, y0, w, y1)
        out = img_path.with_name(f"{img_path.stem}_part{i + 1}.png")
        page.get_pixmap(clip=clip, alpha=False).save(out)
        parts.append(out)
    src.close()
    return parts

# ---------------------------------------------------------------- OCR

# The model writes sara am as two characters (nikhahit U+0E4D + sara aa U+0E32).
# It looks identical but breaks search/spell-check, so compose it into U+0E33,
# keeping any tone mark in front: น + ้ + ํ + า  ->  น + ้ + ำ
_SARA_AM_RE = re.compile("\u0e4d([\u0e48-\u0e4b]?)\u0e32")


def fix_thai(text: str) -> str:
    return _SARA_AM_RE.sub("\\1\u0e33", text)


# ---- confidence marks ---------------------------------------------------
# The model reports how sure it was about every token. Text it was unsure about
# is wrapped in MARK_OPEN/MARK_CLOSE while the raw output is post-processed,
# written to the page markdown as ==text== and shown in Word with a yellow
# highlight, so a proof-reader knows where to look first.
MARK_OPEN, MARK_CLOSE = "", ""
HIGHLIGHT_RE = re.compile(r"==(?=\S)(.+?)(?<=\S)==")
DROP_CLASSES = ("Picture", "Chart")              # not text: photos, maps, graphs
_SPECIAL_TOKEN_RE = re.compile(r"<[^<>\s]+>")
_STOP = set(" \t\r\n&\\{}<>|$" + MARK_OPEN + MARK_CLOSE)
_PAD = 6                                         # characters of context highlighted each side
_NUMBER = set("0123456789๐๑๒๓๔๕๖๗๘๙.,")
_TABULAR_RE = re.compile(r"\\(?:begin|end)\{tabular\}(?:\{[^{}\n]*\})?")
_SPAN_RE = re.compile(r"(\\multi(?:column|row)\{[^{}\n]*\}\{[^{}\n]*\})\{([^{}\n]*)\}")


def _joins_previous(text: str, i: int) -> bool:
    """True if text[i] must stay in the same run as text[i-1] (Thai marks/vowels)."""
    ch = text[i]
    if ch == "ั" or "ิ" <= ch <= "ฺ" or "็" <= ch <= "๎":
        return True                              # vowel/tone mark above or below
    if ch in "าำ" and i and "่" <= text[i - 1] <= "ํ":
        return True                              # sara aa after a tone mark / nikhahit
    return bool(i) and "เ" <= text[i - 1] <= "ไ"   # consonant after a leading vowel


def mark_low_confidence(pieces: list[tuple[str, float]], threshold: float) -> str:
    """Join (token text, probability) pieces, wrapping unsure text in marks."""
    text = "".join(t for t, _ in pieces)
    if threshold <= 0:
        return text
    spans: list[list[int]] = []
    pos = 0
    for tok, prob in pieces:
        end = pos + len(tok)
        body = tok.strip()
        if (prob < threshold and body and not _SPECIAL_TOKEN_RE.fullmatch(body)
                and not any(c in _STOP for c in body)
                and any(c.isalnum() for c in body)                        # not bare punctuation
                and text[max(0, pos - 1):pos + len(tok) - len(tok.lstrip())].strip() != "\\"):  # not a LaTeX command
            a = pos + len(tok) - len(tok.lstrip())
            b = end - (len(tok) - len(tok.rstrip()))
            lo, hi = a, b                        # widen to nearby context, never past a separator
            while lo > 0 and a - lo < _PAD and text[lo - 1] not in _STOP:
                lo -= 1
            while hi < len(text) and hi - b < _PAD and text[hi] not in _STOP:
                hi += 1
            while lo > 0 and text[lo - 1] in _NUMBER and text[lo] in _NUMBER:
                lo -= 1                          # never mark only part of a number
            while hi < len(text) and text[hi] in _NUMBER and text[hi - 1] in _NUMBER:
                hi += 1
            while lo > 0 and text[lo - 1] not in _STOP and _joins_previous(text, lo):
                lo -= 1
            while hi < len(text) and text[hi] not in _STOP and _joins_previous(text, hi):
                hi += 1
            if spans and lo <= spans[-1][1]:     # touches the previous span: merge
                spans[-1] = [min(spans[-1][0], lo), max(spans[-1][1], hi)]
            else:
                spans.append([lo, hi])
        pos = end
    out, last = [], 0
    for lo, hi in spans:
        out += [text[last:lo], MARK_OPEN, text[lo:hi], MARK_CLOSE]
        last = hi
    out.append(text[last:])
    return _fix_table_marks("".join(out))


def _fix_table_marks(text: str) -> str:
    """Keep marks out of LaTeX table structure, where they would break the table.

    An unsure column/row span means the cell may be in the wrong place, so the
    mark moves onto the cell's content instead.
    """
    strip = lambda s: s.replace(MARK_OPEN, "").replace(MARK_CLOSE, "")
    text = _TABULAR_RE.sub(lambda m: strip(m.group(0)), text)

    def span(m):
        if MARK_OPEN not in m.group(1):
            return m.group(0)
        return strip(m.group(1)) + "{" + MARK_OPEN + strip(m.group(2)) + MARK_CLOSE + "}"

    return _SPAN_RE.sub(span, text)


def marks_to_markdown(text: str) -> str:
    return text.replace(MARK_OPEN, "==").replace(MARK_CLOSE, "==")


class VllmBackend:
    """Call the vLLM server and return the raw output with confidence marks."""

    def __init__(self, base_url: str, model: str, threshold: float, drop=DROP_CLASSES):
        import httpx
        self.url = base_url.rstrip("/") + "/chat/completions"
        self.model, self.threshold, self.drop = model, threshold, drop
        self.client = httpx.Client(timeout=1800.0)

    def __call__(self, image, *, prompt, stage, settings):
        import base64
        from dataclasses import asdict
        from io import BytesIO

        stream = BytesIO()
        image.save(stream, format="PNG")
        data = base64.b64encode(stream.getvalue()).decode("ascii")
        payload = {
            "model": self.model,
            "messages": [{"role": "user", "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": "data:image/png;base64," + data}},
            ]}],
            **asdict(settings),
            "skip_special_tokens": False,
            "logprobs": self.threshold > 0,
        }
        reply = self.client.post(self.url, json=payload, headers={"Authorization": "Bearer EMPTY"})
        reply.raise_for_status()
        body = reply.json()
        choice = body["choices"][0]
        raw = choice["message"].get("content") or ""
        tokens = (choice.get("logprobs") or {}).get("content") or []
        pieces = [(t["token"], math.exp(t["logprob"])) for t in tokens]
        if pieces and "".join(t for t, _ in pieces).startswith(raw):
            # tokens line up with the text (they also include the final </s>)
            had_eos = raw.endswith("</s>")
            raw = mark_low_confidence(pieces, self.threshold)
            if raw.endswith("</s>") and not had_eos:
                raw = raw[:-4]
        return {"raw_out": raw,
                "finish_reason": choice.get("finish_reason"), "usage": body.get("usage", {})}


class LocalBackend:
    """Run the model in this process with Transformers (CPU, or GPU if visible)."""

    def __init__(self, model, processor, threshold: float, drop=DROP_CLASSES):
        from threading import Lock
        self.model, self.processor = model, processor
        self.threshold, self.drop = threshold, drop
        self._lock = Lock()

    def __call__(self, image, *, prompt, stage, settings):
        import torch
        from copy import deepcopy

        with self._lock:
            inputs = self.processor(images=[image], text=prompt, return_tensors="pt",
                                    add_special_tokens=False).to(self.model.device)
            config = deepcopy(self.model.generation_config)
            config.max_new_tokens = settings.max_tokens
            config.do_sample = settings.temperature > 0
            config.repetition_penalty = settings.repetition_penalty
            config.top_k = settings.top_k if settings.top_k > 0 else 0
            config.top_p = settings.top_p
            config.temperature = settings.temperature if settings.temperature > 0 else 1.0
            config.return_dict_in_generate = True

            # Record the probability of each token as it is chosen. (Asking generate()
            # for all logits instead keeps a vocabulary-sized tensor per token and
            # exhausts memory on long pages.)
            probs: list[float] = []
            last: list = [None]

            def record(input_ids, scores):
                if last[0] is not None:
                    probs.append(last[0][input_ids[0, -1]].item())
                last[0] = torch.softmax(scores[0].float(), -1)
                return scores

            extra = {}
            if self.threshold > 0:
                from transformers import LogitsProcessorList
                extra["logits_processor"] = LogitsProcessorList([record])
            with torch.inference_mode(), torch.random.fork_rng():
                torch.random.default_generator.manual_seed(settings.seed)
                out = self.model.generate(**inputs, generation_config=config, **extra)
            sequence = out.sequences[0]
            if last[0] is not None:
                probs.append(last[0][sequence[-1]].item())
            last[0] = None
            n_new = len(probs)
            decode = self.processor.tokenizer.decode
            raw = self.processor.batch_decode(out.sequences, skip_special_tokens=False)[0]
            if n_new:
                # probability the model gave each token it produced, and the text each one added
                start = len(sequence) - n_new
                head = decode(sequence[:start], skip_special_tokens=False)
                pieces, prev, carry = [], head, 1.0
                for k in range(n_new):
                    carry = min(carry, probs[k])
                    cur = decode(sequence[:start + k + 1], skip_special_tokens=False)
                    if cur.startswith(prev) and not cur.endswith("�"):
                        pieces.append((cur[len(prev):], carry))
                        prev, carry = cur, 1.0
                if head + "".join(t for t, _ in pieces) == raw:
                    raw = head + mark_low_confidence(pieces, self.threshold)
            prefix = inputs.get("decoder_input_ids")
            if prefix is None:
                prefix = inputs.get("input_ids")
            generated = len(sequence) - (prefix.shape[-1] if prefix is not None else 1)
            return {"raw_out": raw,
                    "finish_reason": "length" if generated >= settings.max_tokens else "stop",
                    "usage": {"completion_tokens": generated}}


def cpu_has_fast_bf16() -> bool:
    """True when the CPU has native bfloat16 math (Intel AMX / AVX512-BF16, AMD Zen 4+).
    Without it bfloat16 on CPU is emulated and slower than float32."""
    try:
        flags = Path("/proc/cpuinfo").read_text(errors="ignore")
    except OSError:
        return False
    return "amx_bf16" in flags or "avx512_bf16" in flags


def load_cpu_backend(model_dir: Path, threads: int | None, threshold: float = 0.8,
                     drop=DROP_CLASSES, dtype: str = "auto", quant: str = "none") -> LocalBackend:
    """Load the model in-process with Transformers (no vLLM server, no GPU needed).

    Much slower than the vLLM server, but runs anywhere. Uses a GPU only if this
    process can see one.

    dtype: "auto" (bfloat16 on GPU, or on a CPU with native bf16; else float32),
           "fp32" or "bf16".
    quant: "none" or "int8" (dynamic int8 for the Linear layers, CPU only: faster,
           but check the output quality on your own documents first).
    """
    import torch
    from transformers import AutoModel, AutoProcessor

    if threads:
        torch.set_num_threads(threads)
    device = "cuda:0" if torch.cuda.is_available() else "cpu"
    if dtype == "auto":
        dtype = "bf16" if device != "cpu" or cpu_has_fast_bf16() else "fp32"
    if quant == "int8" and device == "cpu":
        dtype = "fp32"                               # dynamic int8 quantization needs fp32 weights
    torch_dtype = torch.bfloat16 if dtype == "bf16" else torch.float32
    model = AutoModel.from_pretrained(model_dir, trust_remote_code=True, dtype=torch_dtype,
                                      local_files_only=True).to(device).eval()
    if quant == "int8" and device == "cpu":
        try:
            model = torch.ao.quantization.quantize_dynamic(model, {torch.nn.Linear},
                                                           dtype=torch.qint8)
        except Exception as e:                       # keep running unquantized
            print(f"int8 quantization failed, using fp32: {e}")
            quant = "none"
    processor = AutoProcessor.from_pretrained(model_dir, trust_remote_code=True,
                                              local_files_only=True)
    print(f"โหลดโมเดลแล้ว: {device}, {dtype}{', int8' if quant == 'int8' else ''}, "
          f"{torch.get_num_threads()} threads", flush=True)
    return LocalBackend(model, processor, threshold, drop)


# ---------------------------------------------------------------- several CPU workers
# Generating text one token at a time for one page barely gets faster past ~4-6
# threads, so a single model on a 16-core machine leaves most cores waiting (the
# CPU still shows 100% because idle OpenMP threads spin). Running several copies of
# the model, each with a few threads and its own page, uses the cores far better.
# Each copy is a separate process: the generate() loop is Python and would fight
# over the GIL in threads.

_WORKER_BACKEND = None


def _pool_init(model_dir: str, threads: int, threshold: float, drop, dtype: str, quant: str):
    import torch
    torch.set_num_threads(threads)
    try:
        torch.set_num_interop_threads(1)
    except RuntimeError:
        pass
    global _WORKER_BACKEND
    _WORKER_BACKEND = load_cpu_backend(Path(model_dir), threads, threshold, drop, dtype, quant)


def _pool_ready(_):
    time.sleep(0.5)                                  # keeps this worker busy so the next one starts
    return os.getpid()


def _pool_call(image, prompt, stage, settings):
    return _WORKER_BACKEND(image, prompt=prompt, stage=stage, settings=settings)


def _available_memory_gb() -> float:
    """Memory this process may use: the container limit if there is one, else MemAvailable."""
    limit = None
    for path in ("/sys/fs/cgroup/memory.max", "/sys/fs/cgroup/memory/memory.limit_in_bytes"):
        try:
            raw = Path(path).read_text().strip()
            if raw.isdigit() and int(raw) < 1 << 50:
                limit = int(raw) / 2**30
                break
        except OSError:
            pass
    avail = None
    try:
        for line in Path("/proc/meminfo").read_text().splitlines():
            if line.startswith("MemAvailable:"):
                avail = int(line.split()[1]) / 2**20
    except OSError:
        pass
    vals = [v for v in (limit, avail) if v]
    return min(vals) if vals else 8.0


def plan_cpu_workers(replicas: int | None = None, threads: int | None = None,
                     dtype: str = "auto", quant: str = "none") -> tuple[int, int]:
    """Choose (model copies, threads per copy) for this machine."""
    cpus = len(os.sched_getaffinity(0)) if hasattr(os, "sched_getaffinity") else (os.cpu_count() or 4)
    if replicas and threads:
        return replicas, threads
    if replicas:
        return replicas, max(1, cpus // replicas)
    threads = threads or (4 if cpus >= 8 else cpus)
    bf16 = dtype == "bf16" or (dtype == "auto" and quant != "int8" and cpu_has_fast_bf16())
    # measured on a 16-vCPU Xeon E5-2680 v4 VM: about 1.1 GB resident per fp32 copy
    per_copy_gb = 1.0 if bf16 or quant == "int8" else 1.6     # weights + activations, with headroom
    by_ram = int((_available_memory_gb() - 2) // per_copy_gb)  # keep 2 GB for the rest
    return max(1, min(cpus // threads, by_ram)), threads


class PoolBackend:
    """Several copies of the model in worker processes; call it from several threads
    at once (one page per call) and each call goes to an idle copy."""

    def __init__(self, model_dir: Path, replicas: int, threads: int, threshold: float = 0.8,
                 drop=DROP_CLASSES, dtype: str = "auto", quant: str = "none"):
        import multiprocessing
        from concurrent.futures import ProcessPoolExecutor

        self.drop, self.replicas, self.threads = drop, replicas, threads
        self.pool = ProcessPoolExecutor(
            max_workers=replicas, mp_context=multiprocessing.get_context("spawn"),
            initializer=_pool_init,
            initargs=(str(model_dir), threads, threshold, drop, dtype, quant))
        # start and load every copy now, not when the first pages arrive
        pids = set(f.result() for f in [self.pool.submit(_pool_ready, i) for i in range(replicas)])
        print(f"CPU OCR: {replicas} model copies x {threads} threads ({len(pids)} started)",
              flush=True)

    def __call__(self, image, *, prompt, stage, settings):
        return self.pool.submit(_pool_call, image, prompt, stage, settings).result()

    def close(self):
        self.pool.shutdown(cancel_futures=True)


def load_cpu_pool(model_dir: Path, replicas: int | None = None, threads: int | None = None,
                  threshold: float = 0.8, drop=DROP_CLASSES, dtype: str = "auto",
                  quant: str = "none"):
    """One model in this process if one copy is enough, else a PoolBackend."""
    replicas, threads = plan_cpu_workers(replicas, threads, dtype, quant)
    if replicas == 1:
        backend = load_cpu_backend(model_dir, threads, threshold, drop, dtype, quant)
        backend.replicas = 1
        return backend
    return PoolBackend(model_dir, replicas, threads, threshold, drop, dtype, quant)


def _confident_chars(md: str) -> tuple[int, int]:
    """(characters outside ==marks==, all characters), ignoring tags and spaces."""
    plain = re.sub(r"<[^>]+>", " ", md).replace("[รูปภาพ]", " ")
    sure = HIGHLIGHT_RE.sub("", plain)
    return len(re.sub(r"\s", "", sure)), len(re.sub(r"\s|==", "", plain))


def _looks_sideways(img_path: Path) -> bool:
    """Cheap check for a page whose text lines run vertically (scanned turned 90 degrees).

    Text lines and the gaps between them make the amount of ink vary a lot from
    one pixel row to the next, but hardly at all from column to column. When the
    variation is clearly stronger across columns, the lines run vertically.
    (Measured on test scans: upright pages 0.03-0.3, sideways pages 2-12.)
    """
    import numpy as np
    from PIL import Image

    with Image.open(img_path) as im:
        im = im.convert("L")
        im.thumbnail((1000, 1000))
        ink = np.asarray(im) < 128
    h, w = ink.shape
    ink = ink[h // 12: h - h // 12, w // 12: w - w // 12]      # ignore dark scan edges
    ink = ink[ink.mean(axis=1) < 0.5][:, ink.mean(axis=0) < 0.5]   # and solid bands / rules
    if ink.size == 0 or ink.sum() < 500:
        return False
    rows = ink.sum(axis=1).astype(float)         # one value per row of pixels
    cols = ink.sum(axis=0).astype(float)
    spread = lambda v: v.var() / max(v.mean(), 1e-6) ** 2
    return spread(cols) > 1.5 * spread(rows)


# ---- layout marks -------------------------------------------------------
# The model returns every block with its class and position on the page. They are
# kept in the page markdown as HTML comments (invisible, harmless to edit around):
#     <!--page landscape-->                    first line: page orientation
#     <!--b Text 0.142,0.137,0.911,0.186-->    before a block: class and x0,y0,x1,y1 (0-1)
# A block without numbers is laid out in normal text flow. Pictures, maps and
# charts become a line "[รูปภาพ]", drawn in Word as an empty box of the same size.
PICTURE_TEXT = "[รูปภาพ]"
PAGE_MARK_RE = re.compile(r"<!--page (portrait|landscape)-->\n?")
BLOCK_MARK_RE = re.compile(r"<!--b ([\w-]+)(?: ([\d.]+),([\d.]+),([\d.]+),([\d.]+))?-->\n?")
_HEADING = {"Title": "# ", "Section-header": "## "}


def _document_in_picture(content: str) -> bool:
    """A "picture" that is really a document (a photocopied receipt or letter pasted on
    the page): the model read plenty of text in it and was sure of most of it. Real
    pictures, ID-card copies and handwriting give little text or mostly unsure text.
    (Measured on test scans: receipt copies 89% sure, ID cards and notes 0-8%.)"""
    if not content:
        return False
    sure, total = _confident_chars(marks_to_markdown(content))
    return total >= 80 and sure >= 0.6 * total


def _blocks_to_markdown(blocks: list[dict], size: tuple[int, int], drop) -> tuple[str, dict]:
    """Page markdown with layout marks, plus facts ocr_page needs about the page."""
    w, h = size
    parts = [f"<!--page {'landscape' if w > h else 'portrait'}-->"]
    info = {"dropped": False, "tall": 0, "wide": 0}
    for blk in blocks:
        cls = blk.get("class") or "Text"
        x0, y0, x1, y1 = (min(1.0, max(0.0, float(v))) for v in blk["bbox"])
        x0, x1, y0, y1 = min(x0, x1), max(x0, x1), min(y0, y1), max(y0, y1)   # corners can come in any order
        mark = f"<!--b {cls} {x0:.3f},{y0:.3f},{x1:.3f},{y1:.3f}-->"
        content = (blk.get("content") or "").strip()
        if cls in drop and not _document_in_picture(content):
            info["dropped"] = True
            parts.append(mark + "\n" + PICTURE_TEXT)
            continue
        if not content:
            continue
        if cls not in ("Table", "Picture", "Chart"):
            # text set sideways sits in a box that is taller than it is wide
            info["tall" if (y1 - y0) * h > 1.5 * (x1 - x0) * w else "wide"] += 1
        if cls in DROP_CLASSES:                  # kept on request (--keep-pictures): plain text
            mark = mark.replace(f"<!--b {cls} ", "<!--b Text ")
        parts.append(mark + "\n" + _HEADING.get(cls, "") + content)
    return "\n\n".join(parts), info


def _flow(md: str, landscape: bool | None = None, pictures: bool = True) -> str:
    """Drop block positions (they no longer match the page); optionally set orientation."""
    def unplace(m):
        return f"<!--b {m.group(1)}-->\n" if m.group(1) in DROP_CLASSES else ""
    if not pictures:
        md = re.sub(r"<!--b [^>]*-->\n?" + re.escape(PICTURE_TEXT) + r"\n*", "", md)
    md = BLOCK_MARK_RE.sub(unplace, md)
    if landscape is None:
        return PAGE_MARK_RE.sub("", md)
    return f"<!--page {'landscape' if landscape else 'portrait'}-->\n\n" + PAGE_MARK_RE.sub("", md).lstrip()


def _try_rotations(img_path: Path, read, upright: str, sideways: bool, dropped: bool) -> str:
    """Read the page turned 90 degrees each way; return the text to keep."""
    import difflib
    from PIL import Image

    with Image.open(img_path) as im:
        portrait = im.height >= im.width
    best, best_sure = None, 0
    for angle in (90, 270):
        turned = img_path.with_name(f"{img_path.stem}_rot{angle}.png")
        try:
            with Image.open(img_path) as im:
                im.convert("RGB").rotate(angle, expand=True).save(turned)
            md, _ = read(turned)
        except Exception:                        # a failed extra reading never fails the page
            continue
        finally:
            turned.unlink(missing_ok=True)
        sure, total = _confident_chars(md)
        # only trust a reading the model was mostly sure about (map labels read
        # sideways come back as a few low-confidence fragments)
        if sure >= 60 and sure >= 0.5 * total and sure > best_sure:
            best, best_sure = md, sure
    up_sure, _ = _confident_chars(upright)
    if sideways:
        # The whole page was scanned turned. Prefer the turned reading when it is about
        # as good: its block positions describe the page the right way up. Otherwise
        # keep the text but not the positions, and make the Word page match.
        if best is not None and best_sure >= 0.9 * up_sure:
            return best
        return _flow(upright, landscape=portrait)
    if best is None:
        return upright
    strip = lambda s: re.sub(r"\s|==|<[^>]+>", "", s).replace(PICTURE_TEXT, "")
    same = difflib.SequenceMatcher(None, strip(upright)[:3000], strip(best)[:3000],
                                   autojunk=False).ratio() > 0.5
    if same:                                     # two readings of the same text: keep the surer one
        return _flow(best, landscape=not portrait) if best_sure > 1.15 * up_sure else upright
    # different text (sideways notes inside a plan or map): keep both
    return upright.rstrip() + "\n\n" + _flow(best, pictures=False).strip()


def ocr_page(img_path: Path, backend, rotate: bool = True) -> tuple[str, bool]:
    """Return (markdown, ok). ok is False when any part of the page failed.

    The markdown carries layout marks (see above) and ==text== where the model was
    unsure.

    With rotate, a page that was scanned sideways, or that contains a picture (plans
    and maps often carry sideways notes the model files under "picture"), is also
    read turned 90 degrees each way and the better reading is used or added.
    """
    from nemotron_parse.inference import DocumentParser, GenerationError
    from PIL import Image

    drop = getattr(backend, "drop", DROP_CLASSES)

    with DocumentParser(backend=backend) as parser:
        def read(path) -> tuple[str, dict]:
            with Image.open(path) as im:
                size = im.size
            doc = parser.parse(str(path), mode="single")
            md, info = _blocks_to_markdown(doc.get("blocks") or [], size, drop)
            return fix_thai(marks_to_markdown(md)), info

        try:
            text, info = read(img_path)
            if rotate:
                sideways = info["tall"] >= info["wide"] and _looks_sideways(img_path)
                # Only pages scanned sideways as a whole. Re-reading pages with pictures turned
                # 90 degrees helped OpenThai2.0 find sideways notes in plans; OpenThai2.1 reads
                # those notes as they are, and the turned reading only added repeats and
                # invented text (tested on the kj.1 and CCTV files).
                if sideways:
                    text = _try_rotations(img_path, read, text, sideways, info["dropped"])
            return text, True
        except GenerationError:
            # Page too long for one pass -> read it in two halves (positions would be
            # relative to each half, so this page falls back to normal text flow)
            texts = []
            ok = True
            for part in split_halves(img_path):
                try:
                    texts.append(_flow(read(part)[0]))
                except GenerationError:
                    texts.append(f"> [OCR ไม่สำเร็จ: {part.name}]")
                    ok = False
                finally:
                    part.unlink(missing_ok=True)
            return "\n\n".join(texts), ok

# ---------------------------------------------------------------- Markdown -> DOCX

def make_reference_docx(path: Path) -> None:
    """Create a pandoc reference.docx with Thai fonts on all main styles."""
    from docx import Document
    from docx.oxml.ns import qn
    from docx.shared import Pt

    subprocess.run(
        ["pandoc", "-o", str(path), "--print-default-data-file", "reference.docx"],
        check=True,
    )

    d = Document(str(path))
    sizes = {"Normal": BODY_PT, "Body Text": BODY_PT, "First Paragraph": BODY_PT,
             "Compact": BODY_PT, "Title": 28, "Heading 1": 22, "Heading 2": 20,
             "Heading 3": 18, "Heading 4": 16}
    for style in d.styles:
        if style.type != 1:          # paragraph styles only
            continue
        rpr = style.element.get_or_add_rPr()
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is None:
            fonts = rpr.makeelement(qn("w:rFonts"), {})
            rpr.append(fonts)
        for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            fonts.set(qn(attr), THAI_FONT)
        for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
            fonts.attrib.pop(qn(attr), None)
        if style.name in sizes:
            pt = sizes[style.name]
            style.font.size = Pt(pt)
            szcs = rpr.find(qn("w:szCs"))          # Thai = complex-script size
            if szcs is None:
                szcs = rpr.makeelement(qn("w:szCs"), {})
                rpr.append(szcs)
            szcs.set(qn("w:val"), str(pt * 2))
    # Mark the text as Thai: without it Word only breaks lines at spaces, and Thai
    # sentences (which have none) wrap badly and justify into stretched-out letters.
    defaults = d.styles.element.find(qn("w:docDefaults"))
    if defaults is not None:
        rpr = defaults.find(f"{qn('w:rPrDefault')}/{qn('w:rPr')}")
        if rpr is not None:
            lang = rpr.find(qn("w:lang"))
            if lang is None:
                lang = rpr.makeelement(qn("w:lang"), {})
                rpr.append(lang)
            lang.set(qn("w:bidi"), "th-TH")
    d.save(str(path))


class _TableParser(HTMLParser):
    """Parse an HTML <table> into rows of (text, rowspan, colspan, is_header)."""

    def __init__(self):
        super().__init__()
        self.rows: list[list[tuple[str, int, int, bool]]] = []
        self._cell: list[str] | None = None
        self._attrs: dict = {}
        self._th = False

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self.rows.append([])
        elif tag in ("td", "th"):
            self._cell, self._attrs, self._th = [], dict(attrs), tag == "th"
        elif tag == "br" and self._cell is not None:
            self._cell.append("\n")

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._cell is not None:
            if not self.rows:
                self.rows.append([])
            span = lambda k: max(1, int(str(self._attrs.get(k) or 1).strip() or 1))
            text = re.sub(r"[ \t]+", " ", "".join(self._cell)).strip()
            self.rows[-1].append((text, span("rowspan"), span("colspan"), self._th))
            self._cell = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)


def _layout(rows):
    """Place cells on a grid honouring rowspan/colspan. Returns (cells, n_rows, n_cols)."""
    taken: set[tuple[int, int]] = set()
    placed = []
    for r, row in enumerate(rows):
        c = 0
        for text, rs, cs, th in row:
            while (r, c) in taken:
                c += 1
            placed.append((r, c, rs, cs, text, th))
            for dr in range(rs):
                for dc in range(cs):
                    taken.add((r + dr, c + dc))
            c += cs
    n_rows = max((r for r, _ in taken), default=-1) + 1
    n_cols = max((c for _, c in taken), default=-1) + 1
    return placed, n_rows, n_cols


def _no_borders(cell, sides) -> None:
    from docx.oxml.ns import qn
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = tc_pr.makeelement(qn("w:tcBorders"), {})
        tc_pr.append(borders)
    for side in sides:
        el = borders.find(qn(f"w:{side}"))
        if el is None:
            el = borders.makeelement(qn(f"w:{side}"), {})
            borders.append(el)
        el.set(qn("w:val"), "nil")


def _insert_table(doc, anchor_par, table_html, boxes=None) -> None:
    """Replace the placeholder paragraph with a real Word table.

    table_html may be a tuple of tables that stood side by side on the page (with
    their boxes, left to right). Word cannot put two tables next to each other in
    normal text flow, so they become one table with an empty, borderless column
    between them - which looks the same.
    """
    from docx.enum.text import WD_COLOR_INDEX
    from docx.oxml.ns import qn
    from docx.shared import Pt

    htmls = table_html if isinstance(table_html, tuple) else (table_html,)
    parts = []
    for k, html in enumerate(htmls):
        parser = _TableParser()
        parser.feed(html)
        placed, rows, cols = _layout(parser.rows)
        if placed:
            parts.append((placed, rows, cols, boxes[k] if boxes else None))
    if not parts:
        anchor_par._p.getparent().remove(anchor_par._p)
        return
    n_rows = max(rows for _, rows, _, _ in parts)
    n_cols = sum(cols for _, _, cols, _ in parts) + len(parts) - 1
    table = doc.add_table(rows=n_rows, cols=n_cols)
    table.style = "Table Grid" if "Table Grid" in [st.name for st in doc.styles] else None

    weights = [0.0] * n_cols                     # how much text each column has to hold
    offset = 0
    for placed, rows, cols, _box in parts:
        for r, c, rs, cs, text, th in placed:
            cell = table.cell(r, offset + c)
            if rs > 1 or cs > 1:
                cell = cell.merge(table.cell(min(r + rs, rows) - 1, offset + min(c + cs, cols) - 1))
            par = cell.paragraphs[0]
            for k, piece in enumerate(HIGHLIGHT_RE.split(text)):   # odd pieces were ==unsure==
                if piece:
                    run = par.add_run(piece)
                    if k % 2:
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            for run in par.runs:
                run.bold = th
                run.font.size = Pt(BODY_PT - 2)
                fonts = run._r.get_or_add_rPr().get_or_add_rFonts()
                for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                    fonts.set(qn(attr), THAI_FONT)
                szcs = run._r.rPr.makeelement(qn("w:szCs"), {qn("w:val"): str((BODY_PT - 2) * 2)})
                run._r.rPr.append(szcs)
            if cs == 1:
                longest = max((_width(ln) for ln in text.split("\n")), default=0)
                weights[offset + c] = max(weights[offset + c], longest)
        # rows this part does not have (the neighbour is longer): leave them blank, no lines
        for r in range(rows, n_rows):
            for c in range(cols):
                _no_borders(table.cell(r, offset + c), ("left", "right", "bottom") + (("top",) if r > rows else ()))
        offset += cols + 1

    if len(parts) > 1:
        # spacer columns: no horizontal lines, so the parts look like separate tables
        offset = 0
        for placed, rows, cols, _box in parts[:-1]:
            offset += cols
            for r in range(n_rows):
                _no_borders(table.cell(r, offset), ("top", "bottom"))
            offset += 1
        # column widths: each part as wide as it was on the page, split by its text
        nominal = 450.0                          # points; _fit_table rescales to the real width
        spans = [box[2] - box[0] if box else 1.0 for _, _, _, box in parts]
        gaps = [max(parts[k + 1][3][0] - parts[k][3][2], 0.01) if parts[k][3] and parts[k + 1][3] else 0.04
                for k in range(len(parts) - 1)]
        total = sum(spans) + sum(gaps)
        widths, offset = [0.0] * n_cols, 0
        for k, (_placed, _rows, cols, _box) in enumerate(parts):
            share = [max(w, 3.0) + 2.0 for w in weights[offset:offset + cols]]
            for c in range(cols):
                widths[offset + c] = nominal * spans[k] / total * share[c] / sum(share)
            offset += cols
            if k < len(gaps):
                widths[offset] = nominal * gaps[k] / total
                offset += 1
        table.autofit = False
        for c, width in enumerate(widths):
            table.columns[c].width = Pt(width)
            for cell in table.columns[c].cells:
                cell.width = Pt(width)

    if parts[0][0][0][5]:                        # first row is a header -> repeat on each page
        trpr = table.rows[0]._tr.get_or_add_trPr()
        trpr.append(trpr.makeelement(qn("w:tblHeader"), {qn("w:val"): "true"}))
    anchor_par._p.addnext(table._tbl)            # move table to placeholder position
    anchor_par._p.getparent().remove(anchor_par._p)


def _add_table_grid_style(ref_path: Path) -> None:
    """Give the reference doc a bordered 'Table Grid' table style."""
    from docx import Document
    from docx.oxml import parse_xml
    d = Document(str(ref_path))
    if "Table Grid" in [s.name for s in d.styles]:
        return
    xml = (
        '<w:style xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>'
        '<w:tblPr><w:tblBorders>'
        + "".join(f'<w:{b} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                  for b in ("top", "left", "bottom", "right", "insideH", "insideV"))
        + '</w:tblBorders><w:tblCellMar><w:left w:w="108" w:type="dxa"/>'
        '<w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
    )
    d.styles.element.append(parse_xml(xml))
    d.save(str(ref_path))


TABLE_RE = re.compile(r"<table\b.*?</table>", re.S | re.I)
# error placeholders written by this script when OCR of a page (or half) fails
FAILED_RE = re.compile(r"^> \[(?:หน้า \d+ ผิดพลาด:|OCR ไม่สำเร็จ:)", re.M)


A4 = (11906, 16838)                              # page size in twips (1/20 pt)
CM = 28.35                                       # points per centimetre
# Thai official documents: A4, left 3 cm, right 2 cm, top 1.5-2.5 cm, bottom 2 cm,
# 16 pt type, single line spacing, body paragraphs indented 2.5 cm on the first line.
STD_LEFT, STD_RIGHT, STD_TOP, STD_BOTTOM = 3 * CM, 2 * CM, 2.5 * CM, 2 * CM
STD_FIRST_LINE = 2.5 * CM
# Official documents justify body text ("Thai distributed"). That only looks right when
# Word has its Thai word-breaker; without it Word wraps mid-word and spreads the letters
# of short lines across the page. Left-aligned text is safe everywhere, so it is the default.
JUSTIFY_BODY = False
MIN_MARGIN = 567                                 # 1 cm, in twips


class _PageLayout:
    """Measurements of one source page, taken from its layout marks."""

    def __init__(self, landscape: bool):
        self.landscape = landscape
        w, h = (A4[1], A4[0]) if landscape else A4
        self.w, self.h = w / 20, h / 20          # points
        self.left, self.right, self.top = STD_LEFT, STD_RIGHT, STD_TOP
        self.bottom = STD_BOTTOM
        self.x0, self.x1 = 0.0, 1.0              # content edges as fractions of the page
        self.size = None                         # text size in points, estimated from the scan
        self.dense = False
        self.placed = False

    def measure(self, boxes: list[tuple], sizes: list[float]) -> None:
        if not boxes:
            return
        self.placed = True
        self.x0 = min(b[0] for b in boxes)
        self.x1 = max(max(b[2] for b in boxes), self.x0 + 0.2)
        lo = MIN_MARGIN / 20
        self.left = min(max(self.x0 * self.w, lo), self.w * 0.3)
        self.right = min(max((1 - self.x1) * self.w, lo), self.w * 0.3)
        self.top = min(max(min(b[1] for b in boxes) * self.h - 6, lo), self.h * 0.25)
        self.bottom = 1.5 * CM                   # a little under the standard, as room for error
        # A page filled almost to the bottom has no slack. Word never wraps text exactly
        # like the original (and without its Thai word-breaker it wastes part of every
        # line), so such a page is set slightly tighter to keep it on one page.
        self.dense = max(b[3] for b in boxes) - min(b[1] for b in boxes) > 0.75
        if self.dense:
            self.bottom = 1.0 * CM
        if not self.landscape:                   # close to the standard margins: use them exactly
            if abs(self.left - STD_LEFT) < 0.8 * CM:
                self.left = STD_LEFT
            if abs(self.right - STD_RIGHT) < 0.7 * CM:
                self.right = STD_RIGHT
        if sizes:
            sizes = sorted(sizes)
            self.size = sizes[len(sizes) // 2]

    @property
    def usable(self) -> float:
        return self.w - self.left - self.right

    @property
    def font_pt(self) -> float:                  # never larger than the normal body size
        if self.size is None:
            return BODY_PT
        size = BODY_PT if self.size >= BODY_PT - 2 else self.size    # 14+ is the standard 16 pt
        if self.dense:
            size *= 0.88
        return max(8.0, int(size * 2) / 2)

    def across(self, bbox) -> tuple[float, float]:
        """How far a block starts after the left edge / ends before the right edge (0-1)."""
        span = self.x1 - self.x0
        return (bbox[0] - self.x0) / span, (self.x1 - bbox[2]) / span

    def section_xml(self) -> str:
        orient = ' w:orient="landscape"' if self.landscape else ""
        tw = lambda pt: round(pt * 20)
        return ('<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                f'<w:pgSz w:w="{tw(self.w)}" w:h="{tw(self.h)}"{orient}/>'
                f'<w:pgMar w:top="{tw(self.top)}" w:right="{tw(self.right)}" w:bottom="{tw(self.bottom)}" '
                f'w:left="{tw(self.left)}" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>')


def _set_run_size(run, pt: float) -> None:
    from docx.oxml.ns import qn
    from docx.shared import Pt
    run.font.size = Pt(pt)
    rpr = run._r.get_or_add_rPr()
    szcs = rpr.find(qn("w:szCs"))                # Thai uses the complex-script size
    if szcs is None:
        szcs = rpr.makeelement(qn("w:szCs"), {})
        rpr.append(szcs)
    szcs.set(qn("w:val"), str(round(pt * 2)))


def _set_table_geometry(table_el, width_pt: float, indent_pt: float) -> None:
    from docx.oxml.ns import qn
    tbl_pr = table_el.tblPr
    for tag, val in (("w:tblW", width_pt), ("w:tblInd", indent_pt)):
        el = tbl_pr.find(qn(tag))
        if el is None:
            el = tbl_pr.makeelement(qn(tag), {})
            tbl_pr.append(el)
        el.set(qn("w:w"), str(round(val * 20)))
        el.set(qn("w:type"), "dxa")


def _insert_picture_box(doc, par, bbox, page: _PageLayout) -> None:
    """Put an empty bordered box labelled "รูปภาพ" where a picture was, sized like it."""
    from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor

    usable_h = page.h - page.top - page.bottom
    if bbox and page.placed:
        before, after = page.across(bbox)
        width = max(60.0, min(page.usable, (1 - before - after) * page.usable))
        indent = max(0.0, min(page.usable - width, before * page.usable))
        height = max(40.0, min(usable_h * 0.85, (bbox[3] - bbox[1]) * page.h * 0.9))
    else:
        width, indent, height = page.usable * 0.6, page.usable * 0.2, 150.0

    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid" if "Table Grid" in [s.name for s in doc.styles] else None
    table.autofit = False
    _set_table_geometry(table._tbl, width, indent)
    row, cell = table.rows[0], table.cell(0, 0)
    cell.width = Pt(width)
    row.height = Pt(height)
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    label = cell.paragraphs[0]
    label.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = label.add_run("รูปภาพ")
    run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    fonts = run._r.get_or_add_rPr().get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(attr), THAI_FONT)
    _set_run_size(run, min(BODY_PT, page.font_pt))
    par.text = ""                                # keep an empty paragraph after the box so
    fmt = par.paragraph_format                   # two boxes in a row do not merge into one table
    fmt.space_before = fmt.space_after = Pt(0)
    fmt.line_spacing = Pt(2)
    par._p.addprevious(table._tbl)


def _fit_table(doc, table_el, bbox, page: _PageLayout) -> None:
    """Give a table the width and position it had, and a text size that lets it fit."""
    from docx.oxml.ns import qn
    from docx.shared import Pt
    from docx.table import Table

    before, after = page.across(bbox)
    width = max(80.0, min(page.usable, (1 - before - after) * page.usable))
    if width > 0.6 * page.usable:                # a wide table gets the full width: Word needs
        width, before = page.usable, 0.0         # more room than the original to fit the same text
    _set_table_geometry(table_el, width, max(0.0, min(page.usable - width, before * page.usable)))
    table = Table(table_el, doc._body)
    grid = table_el.findall(f"{qn('w:tblGrid')}/{qn('w:gridCol')}")
    fixed = table_el.tblPr.find(qn("w:tblLayout"))
    if grid and fixed is not None and fixed.get(qn("w:type")) == "fixed":
        scale = width * 20 / max(sum(int(g.get(qn("w:w")) or 0) for g in grid), 1)
        for el in grid + table_el.findall(f".//{qn('w:tcW')}"):
            el.set(qn("w:w"), str(round(int(el.get(qn("w:w")) or 0) * scale)))
    runs = [r for row in table.rows for c in row.cells for par in c.paragraphs for r in par.runs]
    # characters that take up width (Thai vowels/tone marks above and below do not)
    chars = sum(1 for r in runs for ch in r.text
                if not ("ั" == ch or "ิ" <= ch <= "ฺ" or "็" <= ch <= "๎"))
    area = max(width * (bbox[3] - bbox[1]) * page.h, 1.0)    # (the model sometimes returns an empty box)
    # one character needs about 0.45 em x 1.35 em; cell padding, uneven columns and
    # imperfect wrapping waste more than half of a table's area
    size = (0.42 * area / (0.6 * max(chars, 1))) ** 0.5
    size = min(BODY_PT - 2.0, max(7.0, int(size * 2) / 2))
    seen = set()
    for row in table.rows:
        for cell in row.cells:
            if id(cell._tc) in seen:
                continue
            seen.add(id(cell._tc))
            for par in cell.paragraphs:
                par.paragraph_format.space_before = par.paragraph_format.space_after = Pt(0)
                for run in par.runs:
                    _set_run_size(run, size)


def _apply_layout(doc, blocks: list[tuple], pages: list[_PageLayout]) -> None:
    """Use the layout marks: page orientation and margins, text size,
    where each block sits across the page, vertical gaps, table size and picture boxes."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import parse_xml
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor
    from docx.text.paragraph import Paragraph

    body = doc.element.body
    index, cur, first, prev_bottom = 0, None, False, None
    for el in list(body):
        page = pages[index]
        if el.tag == qn("w:tbl"):
            if cur is not None and cur[2] and page.placed:
                _fit_table(doc, el, cur[2], page)
                prev_bottom = cur[2][3]
            cur = None
            continue
        if el.tag != qn("w:p"):
            continue
        par = Paragraph(el, doc._body)
        text = par.text.strip()
        m = re.fullmatch(r"@@B_(\d+)@@", text)
        if m:
            cur, first = blocks[int(m.group(1))], True
            if cur[1] in DROP_CLASSES:
                _insert_picture_box(doc, par, cur[2], page)
                prev_bottom = cur[2][3] if cur[2] else prev_bottom
                cur = None
            else:
                body.remove(el)
            continue
        if text == "@@PAGEBREAK@@":
            # end the page with a section break that carries this page's size and margins
            par.text = ""
            fmt = par.paragraph_format
            fmt.space_before = fmt.space_after = Pt(0)
            fmt.line_spacing = Pt(1)
            el.get_or_add_pPr().append(parse_xml(page.section_xml()))
            index, cur, prev_bottom = index + 1, None, None
            continue
        if cur is None or cur[2] is None or not page.placed:
            continue
        x0, y0, x1, y1 = cur[2]
        fmt = par.paragraph_format
        fmt.space_before = fmt.space_after = Pt(0)
        fmt.line_spacing = 1.0
        size = page.font_pt + (2 if cur[1] == "Title" else 0)
        for run in par.runs:
            _set_run_size(run, size)
            if cur[1] in _HEADING:
                run.font.color.rgb = RGBColor(0, 0, 0)       # pandoc's heading style is blue
        if first:
            first = False
            if prev_bottom is not None:
                gap = ((y0 - prev_bottom) * page.h - 0.2 * page.font_pt) * (0.7 if page.dense else 1)
                if gap > 2:
                    fmt.space_before = Pt(min(round(gap), 150))
            prev_bottom = y1
        if el.find(".//" + qn("w:numPr")) is not None:
            continue                             # list items keep their own indent
        before, after = page.across(cur[2])
        if before > 0.12 and abs(before - after) < 0.08:
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            if before > 0.04:
                fmt.left_indent = Pt(round(before * page.usable))
            # a body paragraph that runs over several lines gets the standard first-line
            # indent (a character is about 0.42 em wide in this typeface)
            capacity = (1 - before) * page.usable / (0.42 * page.font_pt)
            if _width(text) > 1.5 * capacity and el.find(".//" + qn("w:br")) is None:
                if JUSTIFY_BODY:
                    par.alignment = WD_ALIGN_PARAGRAPH.THAI_JUSTIFY
                if not page.landscape and before < 0.04:
                    fmt.first_line_indent = Pt(STD_FIRST_LINE)

    old = body.find(qn("w:sectPr"))              # the last page's section sits at the end of the body
    if old is not None:
        body.remove(old)
    body.append(parse_xml(pages[index].section_xml()))


def _width(text: str) -> float:
    """Rough printed width of a line, in characters (Thai marks above/below take none)."""
    text = re.sub(r"==|<[^>]+>", "", text)
    return sum(0.6 if c == " " else 1.0 for c in text
               if not (c == "ั" or "ิ" <= c <= "ฺ" or "็" <= c <= "๎"))


def _is_thai(ch: str) -> bool:
    return "฀" <= ch <= "๿"


def _reflow(content: str, wide: bool) -> str:
    """Join the lines of a wrapped paragraph so Word can wrap it again.

    The model keeps the scan's line ends. Left as they are, a line that is a
    little too long for the Word page wraps and leaves one word on its own line.
    A line that is clearly shorter than the block's longest line ends a paragraph.
    """
    lines = content.split("\n")
    if not wide or len(lines) < 3 or "<" in content:
        return content
    widths = sorted(_width(ln) for ln in lines if ln.strip())
    full = widths[len(widths) // 2]              # the typical line; one over-long line (two
                                                 # lines the model already joined) must not count
    starts_item = re.compile(r"\s*(?:[-*+•]\s|#|\(?[0-9๐-๙]{1,3}[.)]\s|[0-9๐-๙]+(?:\.[0-9๐-๙]+)+\s)")
    paragraphs, cur = [], lines[0]
    for prev, ln in zip(lines, lines[1:]):
        if not ln.strip() or _width(prev) < 0.8 * full or starts_item.match(ln):
            paragraphs.append(cur)
            cur = ln
        else:
            a, b = cur.rstrip(), ln.lstrip()
            cur = a + ("" if a and b and _is_thai(a[-1]) and _is_thai(b[0]) else " ") + b
    paragraphs.append(cur)
    return "\n\n".join(p for p in paragraphs if p.strip())


_PAGE_NUMBER_RE = re.compile(r"(?m)^(\s*)-(?=\s*(?:==)?[^\s=-]{1,8}(?:==)?\s*-\s*$)")


def _page_blocks(page: str) -> list[tuple]:
    """Split a page's markdown at its layout marks: [(class, bbox or None, content)]."""
    parts = BLOCK_MARK_RE.split(page)
    out = [(None, None, parts[0])] if parts[0].strip() else []
    for k in range(1, len(parts), 6):
        nums = parts[k + 1:k + 5]
        bbox = None
        if nums[0] is not None:                  # the model sometimes gives corners in the wrong order
            x0, y0, x1, y1 = (float(v) for v in nums)
            bbox = (min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))
        cls, content = parts[k], parts[k + 5]
        if cls not in DROP_CLASSES and "<chart" in content:
            cls, content = "Chart", PICTURE_TEXT  # chart data filed under another class
        elif "<t" not in content:                # (tables legitimately repeat empty cells)
            content = _collapse_repeats(content)
        out.append((cls, bbox, content))
    return out


_REPEAT_RE = re.compile(r"(?<!\S)(\S.{3,60}?)(?:\s+\1){4,}(?!\S)", re.S)


def _collapse_repeats(content: str) -> str:
    """The model sometimes gets stuck and writes the same few words hundreds of times.
    Keep one copy, marked unsure, so the reader checks that spot against the scan."""
    bare = content.replace("==", "")
    if not _REPEAT_RE.search(bare):
        return content

    def one(m):
        unit = m.group(1)
        if re.fullmatch(r"[\W_]+", unit):         # dot leaders and rules are fine as they are
            return m.group(0)
        return f"=={unit.strip()} …=="
    return _REPEAT_RE.sub(one, bare)


def _join_side_by_side(page_blocks: list[tuple]) -> list[tuple]:
    """Merge neighbouring Table blocks that sit next to each other on the page into one
    block whose content is a tuple: ((html, ...), (bbox, ...)), left to right."""
    out: list[tuple] = []
    for cls, bbox, content in page_blocks:
        prev = out[-1] if out else None
        if (cls == "Table" and bbox and prev and prev[0] == "Table" and prev[1]
                and len(TABLE_RE.findall(content)) == 1 and not TABLE_RE.sub("", content).strip()):
            group = prev[2] if isinstance(prev[2], tuple) else None
            if group is None and len(TABLE_RE.findall(prev[2])) == 1 and not TABLE_RE.sub("", prev[2]).strip():
                group = ((TABLE_RE.search(prev[2]).group(0),), (prev[1],))
            if group:
                last = group[1][-1]
                overlap = min(bbox[3], last[3]) - max(bbox[1], last[1])
                shorter = min(bbox[3] - bbox[1], last[3] - last[1])
                apart = bbox[0] >= last[2] - 0.01 or bbox[2] <= group[1][0][0] + 0.01
                if shorter > 0 and overlap > 0.6 * shorter and apart:
                    pairs = sorted(zip(group[1] + (bbox,), group[0] + (TABLE_RE.search(content).group(0),)))
                    boxes = tuple(b for b, _ in pairs)
                    union = (min(b[0] for b in boxes), min(b[1] for b in boxes),
                             max(b[2] for b in boxes), max(b[3] for b in boxes))
                    out[-1] = ("Table", union, (tuple(h for _, h in pairs), boxes))
                    continue
        out.append((cls, bbox, content))
    return out


def _text_sizes(page_blocks: list[tuple], layout: "_PageLayout") -> tuple[list[tuple], list[float]]:
    """Every block's box, and text-size estimates (points) from the text blocks.

    A block's area divided by its character count gives the space one character
    takes. Measured on 16 pt Thai official letters, the text size is about 1.32
    times the square root of that.
    """
    boxes = [b for _, b, _ in page_blocks if b]
    sizes = []
    for cls, box, content in page_blocks:
        if not box or cls in ("Table", "Picture", "Chart"):
            continue
        chars = _width(content.replace("\n", "")) if isinstance(content, str) else 0
        if chars >= 60 and box[2] > box[0] and box[3] > box[1]:     # short blocks are dominated by their padding
            area = (box[2] - box[0]) * layout.w * (box[3] - box[1]) * layout.h
            sizes.append(1.32 * (area / chars) ** 0.5)
    return boxes, sizes


def markdown_pages_to_docx(pages_md: list[str], out_docx: Path, work: Path) -> None:
    # 1) Pull HTML tables and layout marks out and leave placeholders
    #    (pandoc 3.1 loses rowspan, and drops HTML comments).
    # 2) Markdown -> HTML -> pandoc -> docx for headings, lists, bold, etc.
    # 3) python-docx swaps table placeholders for real tables with merged cells,
    #    then applies the layout marks.
    tables: list = []                            # html, or ((html, ...), (bbox, ...)) when joined
    blocks: list[tuple] = []                     # (page index, class, bbox or None)
    pages: list[_PageLayout] = []

    def stash(m):
        tables.append(m.group(0))
        return f"\n\n@@TABLE_{len(tables) - 1}@@\n\n"

    body = []
    for i, page in enumerate(pages_md):
        if i:
            body.append("<p>@@PAGEBREAK@@</p>")
        m = PAGE_MARK_RE.search(page)
        layout = _PageLayout(bool(m) and m.group(1) == "landscape")
        page = fix_thai(PAGE_MARK_RE.sub("", page))  # also fixes pages cached by older runs
        page = _PAGE_NUMBER_RE.sub(r"\1\\-", page)   # "- 2 -" is a page number, not a list bullet
        page_blocks = _join_side_by_side(_page_blocks(page))
        layout.measure(*_text_sizes(page_blocks, layout))
        pages.append(layout)
        pieces = []
        for cls, bbox, content in page_blocks:
            if cls is not None:
                blocks.append((i, cls, bbox))
                pieces.append(f"@@B_{len(blocks) - 1}@@")
            if isinstance(content, tuple):       # tables side by side -> one joined table
                tables.append(content)
                pieces.append(f"@@TABLE_{len(tables) - 1}@@")
                continue
            if cls is not None:
                if bbox and cls not in ("Table",) + DROP_CLASSES:
                    wide = sum(layout.across(bbox)) < 0.35       # spans most of the text width
                    content = _reflow(content.strip("\n"), wide)
            pieces.append(content.replace(PICTURE_TEXT, ""))
        page = TABLE_RE.sub(stash, "\n\n".join(pieces))
        page = HIGHLIGHT_RE.sub(r"<mark>\1</mark>", page)     # ==unsure== -> yellow highlight
        # nl2br: a line break inside a block is a real line break on the page
        body.append(md_lib.markdown(page, extensions=["extra", "sane_lists", "nl2br"]))
    html = ("<!DOCTYPE html><html><head><meta charset='utf-8'></head><body>"
            + "\n".join(body) + "</body></html>")
    html_path = work / "doc.html"
    html_path.write_text(html, encoding="utf-8")

    ref = work / "reference.docx"
    make_reference_docx(ref)
    _add_table_grid_style(ref)
    subprocess.run(
        ["pandoc", str(html_path), "-f", "html", "-t", "docx",
         "--reference-doc", str(ref), "-o", str(out_docx)],
        check=True,
    )

    from docx import Document
    doc = Document(str(out_docx))
    for par in list(doc.paragraphs):
        m = re.fullmatch(r"\s*@@TABLE_(\d+)@@\s*", par.text)
        if m:
            item = tables[int(m.group(1))]
            if isinstance(item, tuple):
                _insert_table(doc, par, item[0], item[1])
            else:
                _insert_table(doc, par, item)
    _apply_layout(doc, blocks, pages or [_PageLayout(False)])
    doc.save(str(out_docx))

# ---------------------------------------------------------------- main

def parse_range(spec: str | None, n: int) -> list[int]:
    if not spec:
        return list(range(n))
    out = []
    for part in spec.split(","):
        a, _, b = part.partition("-")
        lo, hi = int(a), int(b or a)
        out.extend(range(lo - 1, min(hi, n)))
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdf", type=Path)
    ap.add_argument("-o", "--output", type=Path)
    ap.add_argument("--pages", help="e.g. 1-10,15")
    ap.add_argument("--workers", type=int, default=4, help="concurrent pages sent to vLLM")
    ap.add_argument("--base-url", default="http://localhost:8000/v1")
    ap.add_argument("--model", default="nemotron-parse")
    ap.add_argument("--local", action="store_true",
                    help="run the model inside this process with Transformers instead of "
                         "calling a vLLM server (works on CPU; slow)")
    ap.add_argument("--model-dir", type=Path, default=Path(__file__).resolve().parent / "model",
                    help="model folder used by --local")
    ap.add_argument("--threads", type=int,
                    help="CPU threads per model copy for --local (default: 4)")
    ap.add_argument("--replicas", type=int,
                    help="model copies for --local, each reads its own page "
                         "(default: CPUs / threads, limited by free RAM)")
    ap.add_argument("--dtype", choices=["auto", "fp32", "bf16"], default="auto",
                    help="--local precision (auto = bf16 only on CPUs with native bf16)")
    ap.add_argument("--quant", choices=["none", "int8"], default="none",
                    help="--local: int8 dynamic quantization (faster; check quality first)")
    ap.add_argument("--confidence", type=float, default=0.8, metavar="0-1",
                    help="highlight text the model was less sure about than this "
                         "(default 0.8; higher marks more, 0 turns highlighting off)")
    ap.add_argument("--keep-pictures", action="store_true",
                    help="keep what the model reads inside photos, maps and charts")
    ap.add_argument("--no-rotate", action="store_true",
                    help="do not re-read sideways pages and plan/map pages turned 90 degrees")
    ap.add_argument("--from-md", type=Path,
                    help="skip OCR: build DOCX from a folder of page_XXXXX.md files")
    args = ap.parse_args()

    out = args.output or args.pdf.with_suffix(".docx")
    cache = out.with_suffix(".pages")        # per-page markdown, lets you resume
    cache.mkdir(exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)

        if args.from_md:
            md_files = sorted(args.from_md.glob("page_*.md"))
            pages_md = [p.read_text(encoding="utf-8") for p in md_files]
        else:
            doc = pymupdf.open(args.pdf)
            if doc.needs_pass:
                sys.exit("PDF ติดรหัสผ่าน")
            indices = parse_range(args.pages, doc.page_count)
            print(f"{args.pdf.name}: {doc.page_count} หน้า, จะ OCR {len(indices)} หน้า")

            results: dict[int, str] = {}
            todo = []
            for i in indices:
                cached = cache / f"page_{i + 1:05d}.md"
                text = cached.read_text(encoding="utf-8") if cached.exists() else None
                # older runs cached error placeholders; treat those as not done
                if text is not None and not FAILED_RE.search(text):
                    results[i] = text
                else:
                    todo.append(i)
            if len(todo) < len(indices):
                print(f"ข้าม {len(indices) - len(todo)} หน้าที่ทำไว้แล้ว")

            drop = () if args.keep_pictures else DROP_CLASSES
            if args.local and todo:
                backend = load_cpu_pool(args.model_dir, args.replicas, args.threads,
                                        args.confidence, drop, args.dtype, args.quant)
                args.workers = backend.replicas      # one page per model copy
            else:
                backend = VllmBackend(args.base_url, args.model, args.confidence, drop)

            t0 = time.time()
            done = 0
            failed: list[int] = []
            pending: dict = {}
            queue = iter(todo)
            with ThreadPoolExecutor(max_workers=args.workers) as pool:
                while True:
                    # keep only a small window of rendered pages on disk (100MB PDFs!)
                    while len(pending) < args.workers * 2:
                        i = next(queue, None)
                        if i is None:
                            break
                        img = render_page(doc, i, work)
                        pending[pool.submit(ocr_page, img, backend, not args.no_rotate)] = (i, img)
                    if not pending:
                        break
                    finished, _ = wait(pending, return_when=FIRST_COMPLETED)
                    for fut in finished:
                        i, img = pending.pop(fut)
                        try:
                            text, ok = fut.result()
                            text = fix_thai(text)
                        except Exception as e:       # keep going on single-page failure
                            text, ok = f"> [หน้า {i + 1} ผิดพลาด: {e}]", False
                        results[i] = text
                        cached = cache / f"page_{i + 1:05d}.md"
                        if ok:
                            cached.write_text(text, encoding="utf-8")
                        else:                        # don't cache failures -> retried on next run
                            cached.unlink(missing_ok=True)
                            failed.append(i + 1)
                        img.unlink(missing_ok=True)
                        done += 1
                        rate = (time.time() - t0) / done
                        mark = "" if ok else "  ล้มเหลว"
                        print(f"  [{done}/{len(todo)}] หน้า {i + 1}  ~{rate:.1f}s/หน้า{mark}", flush=True)
            doc.close()
            if failed:
                print(f"OCR ไม่สำเร็จ {len(failed)} หน้า: {sorted(failed)} "
                      "(รันคำสั่งเดิมซ้ำเพื่อลองใหม่เฉพาะหน้าเหล่านี้)")
            pages_md = [results[i] for i in indices]

        markdown_pages_to_docx(pages_md, out, work)
    print(f"เสร็จ: {out}")


if __name__ == "__main__":
    main()
