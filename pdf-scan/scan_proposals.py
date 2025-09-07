#!/usr/bin/env python3
import os, re, io, subprocess, sys
from pathlib import Path
from typing import Dict, List
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import pandas as pd
import typer
from typing import Optional

app = typer.Typer(add_completion=False)

# --------- Defaults (overridable by CSV) ----------
DEFAULT_SYNONYMS = {
    "Title":        [r"^\s*(proposed\s+topic|title of research|project title|title)\b"],
    "Background":   [r"^\s*(background|overview|introduction)\b"],
    "Methodology":  [r"^\s*(methodology|research methods|materials and methods)\b"],
    "OutcomesValue":[r"^\s*(outcomes?\b|value\b|findings\b|analysis)\b"],
    "References":   [r"^\s*(references|bibliography|works cited)\b"],
}
# OCR config
TESS_LANG = "eng"          # add "chi_sim+eng" if you need Simplified Chinese & English (must have traineddata)
TESS_CFG  = r"--oem 3 --psm 6"  # LSTM, assume blocks of text
DPI       = 150           # balance speed/accuracy

def load_synonyms(csv_path: Path|None) -> Dict[str, List[re.Pattern]]:
    if csv_path and csv_path.exists():
        df = pd.read_csv(csv_path)
        d: Dict[str, List[str]] = {}
        for _, row in df.iterrows():
            d.setdefault(str(row["Section"]).strip(), []).append(str(row["KeywordRegex"]).strip())
        compiled = {k: [re.compile(p, re.I | re.M) for p in v] for k, v in d.items()}
        return compiled
    return {k: [re.compile(p, re.I | re.M) for p in v] for k, v in DEFAULT_SYNONYMS.items()}

def has_text_layer(pdf: Path) -> bool:
    try:
        with fitz.open(pdf) as doc:
            if doc.page_count == 0:
                return False
            text = doc[0].get_text("text")
            return bool(text and text.strip())
    except Exception:
        return False

def ensure_ocr(pdf: Path, ocr_dir: Path|None) -> Path:
    """
    If the PDF has no text layer and ocrmypdf is available, produce an OCR'ed PDF into ocr_dir.
    Otherwise fall back to page raster + pytesseract on the fly.
    """
    if has_text_layer(pdf):
        return pdf
    # try ocrmypdf if installed
    if shutil_which("ocrmypdf"):
        ocr_dir.mkdir(parents=True, exist_ok=True)
        out = ocr_dir / (pdf.stem + ".ocr.pdf")
        if out.exists():
            return out
        try:
            subprocess.run(
                ["ocrmypdf", "--skip-text", "--rotate-pages", "--optimize", "1", str(pdf), str(out)],
                check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
            )
            if out.exists():
                return out
        except Exception:
            pass
    # fallback: return original and we will OCR pages with pytesseract when extracting
    return pdf

def shutil_which(cmd: str) -> str|None:
    from shutil import which
    return which(cmd)

def pdf_to_text(pdf: Path, pages_limit: int|None=None, force_page_ocr: bool=False) -> str:
    """
    Extract text from a PDF; if no text layer (or force_page_ocr), raster pages and run pytesseract.
    """
    out_text_parts: List[str] = []
    with fitz.open(pdf) as doc:
        page_count = doc.page_count
        N = min(page_count, pages_limit) if pages_limit else page_count
        for i in range(N):
            page = doc[i]
            t = page.get_text("text") if not force_page_ocr else ""
            if not t or not t.strip():
                # OCR this page
                mat = fitz.Matrix(DPI/72, DPI/72)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                txt = pytesseract.image_to_string(img, lang=TESS_LANG, config=TESS_CFG)
                out_text_parts.append(txt)
            else:
                out_text_parts.append(t)
    return "\n".join(out_text_parts)

def find_matches(text: str, compiled_synonyms: Dict[str, List[re.Pattern]]):
    found_map: Dict[str, bool] = {}
    which_map: Dict[str, List[str]] = {}
    for section, regs in compiled_synonyms.items():
        labels = set()
        hit = False
        for rg in regs:
            for m in rg.finditer(text):
                label = m.group(1) if m.lastindex else m.group(0)
                labels.add(label.strip())
                hit = True
        found_map[section] = hit
        which_map[section] = sorted(labels)
    return found_map, which_map

def detect_missing_pages(pdf: Path) -> str:
    """
    Heuristic: check printed page numbers continuity and odd/even-only pattern.
    Returns a short remark string if anything suspicious is found; else ''.
    """
    re_num = re.compile(r"\bpage\s+(\d+)\b(?:\s*of\s*(\d+))?", re.I)
    re_lone = re.compile(r"^\s*(\d{1,3})\s*$")
    nums = []
    with fitz.open(pdf) as doc:
        for i in range(doc.page_count):
            text = doc[i].get_text("text")
            cand = []
            for line in text.splitlines():
                m = re_num.search(line)
                if m:
                    cand.append(int(m.group(1)))
                    continue
                m2 = re_lone.match(line)
                if m2:
                    cand.append(int(m2.group(1)))
            nums.append(min(cand, key=lambda x: abs(x-(i+1))) if cand else None)
    present = [n for n in nums if n is not None]
    if not present:
        return ""
    s = sorted(set(present))
    gaps = [n for n in range(s[0], s[-1]+1) if n not in s]
    odd_only = present and all(n % 2 == 1 for n in present)
    even_only = present and all(n % 2 == 0 for n in present)
    bits = []
    if gaps:
        bits.append(f"Page numbering gaps: {gaps[:10]}{'…' if len(gaps)>10 else ''}")
    if odd_only or even_only:
        bits.append("Odd pages only" if odd_only else "Even pages only")
    return "; ".join(bits)

@app.command()
def scan(
    input_dir: Path = typer.Option(..., "--input-dir", "-i",
                                   help="Folder containing PDFs (recursively scanned)"),
    out_csv: Path = typer.Option("scan_results.csv", "--out-csv", "-o",
                                 help="Output CSV path"),
    synonyms_csv: Optional[Path] = typer.Option(None, "--synonyms-csv", "-s",
                                                help="Optional synonyms CSV (Section,KeywordRegex)"),
    ocr_cache_dir: Path = typer.Option("ocr_out", "--ocr-cache-dir",
                                       help="Where to put OCR'ed PDFs if ocrmypdf is available"),
    pages_limit: int = typer.Option(0, "--pages-limit",
                                    help="0 = all pages; otherwise limit pages per PDF for speed"),
    check_pages: bool = typer.Option(True, "--check-pages/--no-check-pages",
                                     help="Try to detect missing pages and note in REMARKS"),
):
    compiled_syn = load_synonyms(synonyms_csv)
    rows = []

    pdfs = []
    for root, _, files in os.walk(input_dir):
        for f in files:
            if f.lower().endswith(".pdf"):
                pdfs.append(Path(root) / f)
    pdfs.sort()

    for idx, pdf in enumerate(pdfs, 1):
        # typer.echo(f"Scanning {pdf.name} ({idx+1}/{len(pdfs)})...")
        typer.echo(f"[{idx}/{len(pdfs)}] Scanning {pdf.name} ...")
        try:
            # If no text, try to OCR via ocrmypdf; otherwise we'll fallback to per-page OCR later.
            prepared = ensure_ocr(pdf, Path(ocr_cache_dir))
            # If prepared is still non-searchable, pdf_to_text() will OCR pages on the fly.
            text = pdf_to_text(prepared, pages_limit if pages_limit>0 else None)

            # section matches
            found_map, which_map = find_matches(text, compiled_syn)

            # per-section rows
            for sec in compiled_syn.keys():
                mk = which_map.get(sec, [])
                yesno = "YES" if found_map.get(sec, False) else "NO"
                remark_bits = []
                if yesno == "NO":
                    remark_bits.append("No heading match detected")
                if check_pages:
                    page_flag = detect_missing_pages(prepared if prepared.exists() else pdf)
                    if page_flag:
                        remark_bits.append(page_flag)
                rows.append({
                    "PDF FILE NAME": pdf.name,
                    "SECTION_SYNONYMS": sec,
                    "PDF 中有哪個KEYWORD 中了": ", ".join(mk),
                    "結果 YES/NO": yesno,
                    "REMARKS": "; ".join(remark_bits)
                })
        except Exception as e:
            for sec in compiled_syn.keys():
                rows.append({
                    "PDF FILE NAME": pdf.name,
                    "SECTION_SYNONYMS": sec,
                    "PDF 中有哪個KEYWORD 中了": "",
                    "結果 YES/NO": "ERROR",
                    "REMARKS": f"{type(e).__name__}: {e}"
                })

    pd.DataFrame(rows).to_csv(out_csv, index=False)
    typer.secho(f"Wrote {out_csv} with {len(rows)} rows (files scanned: {len(pdfs)})", fg=typer.colors.GREEN)

if __name__ == "__main__":
    app()
