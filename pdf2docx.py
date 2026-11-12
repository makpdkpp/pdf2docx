#!/usr/bin/env python3
"""
PDF -> OpenThai2.0-OCR-Markdown (via vLLM) -> MS Word (.docx)

Usage:
    python pdf2docx.py input.pdf -o output.docx
    python pdf2docx.py input.pdf -o output.docx --pages 1-20 --workers 4

Requirements:
    - vLLM server already running (see README.md)
    - pip install pymupdf python-docx markdown nemotron-parse
    - pandoc >= 3.0 on PATH
"""
from __future__ import annotations

import argparse
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
THAI_FONT = "TH Sarabun New"
BODY_PT = 16

# ---------------------------------------------------------------- PDF -> image

def render_page(doc: pymupdf.Document, index: int, out_dir: Path) -> Path:
    """Render one page so it fits inside 1664x2048 px, keeping aspect ratio."""
    page = doc[index]
    w, h = page.rect.width, page.rect.height        # in points
    zoom = min(MODEL_W / w, MODEL_H / h)
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
    path = out_dir / f"page_{index + 1:05d}.png"
    pix.save(path)
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

def ocr_page(img_path: Path, base_url: str, model: str) -> str:
    from nemotron_parse.inference import DocumentParser, GenerationError

    with DocumentParser(base_url=base_url, model=model) as parser:
        try:
            return parser.parse(str(img_path), mode="single")["markdown"]
        except GenerationError:
            # Page too long for one pass -> read it in two halves
            texts = []
            for part in split_halves(img_path):
                try:
                    texts.append(parser.parse(str(part), mode="single")["markdown"])
                except GenerationError:
                    texts.append(f"> [OCR ไม่สำเร็จ: {part.name}]")
            return "\n\n".join(texts)

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


def _insert_table(doc, anchor_par, table_html: str) -> None:
    from docx.oxml.ns import qn
    from docx.shared import Pt

    p = _TableParser()
    p.feed(table_html)
    placed, n_rows, n_cols = _layout(p.rows)
    if not placed:
        anchor_par._p.getparent().remove(anchor_par._p)
        return
    table = doc.add_table(rows=n_rows, cols=n_cols)
    table.style = "Table Grid" if "Table Grid" in [s.name for s in doc.styles] else None
    for r, c, rs, cs, text, th in placed:
        cell = table.cell(r, c)
        if rs > 1 or cs > 1:
            cell = cell.merge(table.cell(min(r + rs, n_rows) - 1, min(c + cs, n_cols) - 1))
        cell.text = text
        for run in cell.paragraphs[0].runs:
            run.bold = th
            run.font.size = Pt(BODY_PT - 2)
            fonts = run._r.get_or_add_rPr().get_or_add_rFonts()
            for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                fonts.set(qn(attr), THAI_FONT)
            szcs = run._r.rPr.makeelement(qn("w:szCs"), {qn("w:val"): str((BODY_PT - 2) * 2)})
            run._r.rPr.append(szcs)
    if placed and placed[0][5]:                      # first row is header -> repeat on each page
        trpr = table.rows[0]._tr.get_or_add_trPr()
        trpr.append(trpr.makeelement(qn("w:tblHeader"), {qn("w:val"): "true"}))
    anchor_par._p.addnext(table._tbl)                # move table to placeholder position
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


def markdown_pages_to_docx(pages_md: list[str], out_docx: Path, work: Path) -> None:
    # 1) Pull HTML tables out and leave placeholders (pandoc 3.1 loses rowspan).
    # 2) Markdown -> HTML -> pandoc -> docx for headings, lists, bold, etc.
    # 3) python-docx swaps each placeholder for a real table with merged cells.
    tables: list[str] = []

    def stash(m):
        tables.append(m.group(0))
        return f"\n\n@@TABLE_{len(tables) - 1}@@\n\n"

    body = []
    for i, page in enumerate(pages_md):
        if i:
            body.append("<p>@@PAGEBREAK@@</p>")
        page = TABLE_RE.sub(stash, page)
        body.append(md_lib.markdown(page, extensions=["extra", "sane_lists"]))
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
    from docx.enum.text import WD_BREAK
    doc = Document(str(out_docx))
    for par in list(doc.paragraphs):
        m = re.fullmatch(r"\s*@@TABLE_(\d+)@@\s*", par.text)
        if m:
            _insert_table(doc, par, tables[int(m.group(1))])
        elif par.text.strip() == "@@PAGEBREAK@@":
            par.text = ""
            par.add_run().add_break(WD_BREAK.PAGE)
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
                if cached.exists():
                    results[i] = cached.read_text(encoding="utf-8")
                else:
                    todo.append(i)
            if len(todo) < len(indices):
                print(f"ข้าม {len(indices) - len(todo)} หน้าที่ทำไว้แล้ว")

            t0 = time.time()
            done = 0
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
                        pending[pool.submit(ocr_page, img, args.base_url, args.model)] = (i, img)
                    if not pending:
                        break
                    finished, _ = wait(pending, return_when=FIRST_COMPLETED)
                    for fut in finished:
                        i, img = pending.pop(fut)
                        try:
                            text = fut.result()
                        except Exception as e:       # keep going on single-page failure
                            text = f"> [หน้า {i + 1} ผิดพลาด: {e}]"
                        results[i] = text
                        (cache / f"page_{i + 1:05d}.md").write_text(text, encoding="utf-8")
                        img.unlink(missing_ok=True)
                        done += 1
                        rate = (time.time() - t0) / done
                        print(f"  [{done}/{len(todo)}] หน้า {i + 1}  ~{rate:.1f}s/หน้า", flush=True)
            doc.close()
            pages_md = [results[i] for i in indices]

        markdown_pages_to_docx(pages_md, out, work)
    print(f"เสร็จ: {out}")


if __name__ == "__main__":
    main()
