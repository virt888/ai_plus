#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import io
import subprocess
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import pandas as pd
import typer

app = typer.Typer(add_completion=False)

# ===============================
# OCR / Text extraction config
# ===============================
TESS_LANG = "eng"          # e.g., "chi_tra+eng" if you also want Traditional Chinese
TESS_CFG  = r"--oem 3 --psm 6"
DPI       = 150

# ===============================
# Section synonyms (default fallback if no synonyms.csv)
# ===============================
DEFAULT_SECTION_SYNONYMS = {
    "Title": [
        r"^\s*(proposed\s+topic|title of research|project title|title)\b",
    ],
    "Background": [
        r"^\s*(background|overview|introduction)\b",
    ],
    "Methodology": [
        r"^\s*(methodology|research methods|materials and methods)\b",
    ],
    # OutcomesValue: tolerant – will match “Expected Outcomes”, numbered headings, etc.
    "OutcomesValue": [
        r"^\s*(?:\d+[\.\)]\s*)?(?:project\s+)?(?:expected\s+|anticipated\s+)?(?:outcomes?|results?|deliverables|outputs?|impact|value|significance|importance|contribution|novelty|innovation|findings)\b",
    ],
    "References": [
        r"^\s*(list of references?|reference list|references?|bibliography|works cited)\b",
    ],
}

# ===============================
# Optional built-in defaults you may want in addition to Issues.xlsx
# (will be used only if --with-defaults is passed)
# ===============================
# NOTE: The default issue rules use the "MISSING_*" prefix for consistency with
# the user-provided Issues.xlsx.  Any legacy MUST_HAVE_* rule IDs are mapped
# to MISSING_* later when the rules are loaded.  See the logic in the
# `main()` function where `rename_map` is applied to each rule.
DEFAULT_ISSUE_RULES = [
    {"IssueID": "MISSING_PAGES", "IssueName": "Missing / odd-even-only pages",
     "RuleType": "missing_pages", "Regex": "", "MustBePresent": True, "Scope": "any"},
    {"IssueID": "MISS_SEC_Title", "IssueName": "Missing Title section",
     "RuleType": "missing_section", "Section": "Title", "Regex": "", "MustBePresent": True, "Scope": "any"},
    {"IssueID": "MISS_SEC_Background", "IssueName": "Missing Background section",
     "RuleType": "missing_section", "Section": "Background", "Regex": "", "MustBePresent": True, "Scope": "any"},
    {"IssueID": "MISS_SEC_Methodology", "IssueName": "Missing Methodology section",
     "RuleType": "missing_section", "Section": "Methodology", "Regex": "", "MustBePresent": True, "Scope": "any"},
    {"IssueID": "MISS_SEC_OutcomesValue", "IssueName": "Missing Outcomes/Value section",
     "RuleType": "missing_section", "Section": "OutcomesValue", "Regex": "", "MustBePresent": True, "Scope": "any"},
    {"IssueID": "MISS_SEC_References", "IssueName": "Missing References section",
     "RuleType": "missing_section", "Section": "References", "Regex": "", "MustBePresent": True, "Scope": "any"},
    {"IssueID": "MISSING_TIMELINE", "IssueName": "Timeline/schedule not found",
     "RuleType": "keyword",
     "Regex": r"\b(timeline|schedule|gantt|milestone|work\s*plan|workplan|time\s*line)\b",
     "MustBePresent": True, "Scope": "any"},
    {"IssueID": "MISSING_BUDGET", "IssueName": "Budget/cost not found",
     "RuleType": "keyword",
     "Regex": r"\b(budget|costs?|expenditure|expenses?|funding)\b",
     "MustBePresent": True, "Scope": "any"},
]

# ===============================
# Helpers
# ===============================
def _which(cmd: str) -> Optional[str]:
    from shutil import which
    return which(cmd)

def _norm_key(s: str) -> str:
    """lowercase + remove non-alphanumerics, e.g. 'Outcomes & Value' -> 'outcomesvalue'."""
    return re.sub(r"[^a-z0-9]+", "", s.lower())

def build_section_key_map(section_names: List[str]) -> Dict[str, str]:
    """Map normalized aliases -> canonical section key used by synonyms."""
    m = {_norm_key(k): k for k in section_names}
    # Common human variants that should map to OutcomesValue
    for alias in [
        "outcomesvalue", "outcomesandvalue", "outcomes", "expectedoutcomes",
        "keyoutcomes", "outcomesimpact", "results", "deliverables"
    ]:
        m.setdefault(alias, "OutcomesValue")
    return m

def compile_section_synonyms(mapping: Dict[str, List[str]]) -> Dict[str, List[re.Pattern]]:
    return {k: [re.compile(p, re.I | re.M) for p in v] for k, v in mapping.items()}

def load_synonyms_csv(path: Optional[Path]) -> Dict[str, List[str]]:
    """Load {Section: [regex,...]} from synonyms.csv; fallback to defaults if not provided."""
    if not path or not path.exists():
        return {k: v[:] for k, v in DEFAULT_SECTION_SYNONYMS.items()}

    sdf = pd.read_csv(path)
    sdf = sdf.rename(columns=lambda c: str(c).strip())
    if "Section" not in sdf.columns or "KeywordRegex" not in sdf.columns:
        raise ValueError("synonyms.csv must have columns: Section, KeywordRegex")

    sdf["Section"] = sdf["Section"].astype(str).fillna("").str.strip()
    sdf["KeywordRegex"] = sdf["KeywordRegex"].astype(str).fillna("").str.strip()
    sdf = sdf[(sdf["Section"] != "") & (sdf["KeywordRegex"] != "")]

    syn_map: Dict[str, List[str]] = {}
    for _, row in sdf.iterrows():
        sec = row["Section"]
        pat = row["KeywordRegex"]
        try:
            re.compile(pat)
        except re.error as e:
            typer.echo(f"[WARN] Invalid regex for section '{sec}': {pat} ({e})")
            continue
        syn_map.setdefault(sec, []).append(pat)

    if not syn_map:
        raise ValueError("No valid (Section, KeywordRegex) pairs found in synonyms.csv")

    return syn_map

def pdf_has_text_layer(pdf: Path) -> bool:
    try:
        with fitz.open(pdf) as doc:
            if doc.page_count == 0:
                return False
            t = doc[0].get_text("text")
            return bool(t and t.strip())
    except Exception:
        return False

def ensure_ocr_pdf(pdf: Path, ocr_cache_dir: Path) -> Path:
    """If no text layer and ocrmypdf exists, produce OCR'ed PDF; else return original."""
    if pdf_has_text_layer(pdf):
        return pdf
    if _which("ocrmypdf"):
        ocr_cache_dir.mkdir(parents=True, exist_ok=True)
        out = ocr_cache_dir / f"{pdf.stem}.ocr.pdf"
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
    return pdf

def page_to_text_via_ocr(page) -> str:
    mat = fitz.Matrix(DPI/72, DPI/72)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    return pytesseract.image_to_string(img, lang=TESS_LANG, config=TESS_CFG)

def pdf_to_text(pdf: Path, pages_limit: Optional[int]) -> str:
    parts: List[str] = []
    with fitz.open(pdf) as doc:
        N = min(doc.page_count, pages_limit) if pages_limit else doc.page_count
        for i in range(N):
            page = doc[i]
            t = page.get_text("text")
            if not t or not t.strip():
                t = page_to_text_via_ocr(page)
            parts.append(t)
    return "\n".join(parts)

# --- helpers ---------------------------------------------------------------

def _roman_to_int(s: str) -> Optional[int]:
    m = {'I':1,'V':5,'X':10,'L':50,'C':100,'D':500,'M':1000}
    s = (s or "").upper().strip()
    if not re.fullmatch(r"[IVXLCDM]{1,7}", s):
        return None
    val, prev = 0, 0
    for ch in s:
        cur = m.get(ch, 0)
        if cur == 0:
            return None
        val += cur - 2*prev if cur > prev else cur
        prev = cur
    return val if 1 <= val <= 4000 else None

def _ocr_strip_numbers(page, band: float = 0.22, where: str = "bottom") -> list[int]:
    """OCR a horizontal strip (top/bottom) and return candidate integers."""
    try:
        band = max(0.08, min(0.30, float(band)))
        h, w = page.rect.height, page.rect.width
        clip = fitz.Rect(0, h*(1-band), w, h) if where == "bottom" else fitz.Rect(0, 0, w, h*band)
        mat = fitz.Matrix(DPI/72, DPI/72)
        pix = page.get_pixmap(matrix=mat, clip=clip, alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        ocr_cfg = TESS_CFG + " -c tessedit_char_whitelist=0123456789IVXLCDM"
        raw = pytesseract.image_to_string(img, lang=TESS_LANG, config=ocr_cfg)
        tokens = re.findall(r"[IVXLCDM]+|\d{1,4}", raw, flags=re.I)
        out: list[int] = []
        for t in tokens:
            if t.isdigit():
                out.append(int(t))
            else:
                rv = _roman_to_int(t)
                if rv is not None:
                    out.append(rv)
        return out
    except Exception:
        return []

def _import_pypdf():
    """
    Try to import pypdf (preferred) or PyPDF2 as a fallback.
    Returns the module or None if neither is available.
    """
    try:
        import pypdf as _p
        return _p
    except Exception:
        pass
    try:
        import PyPDF2 as _p
        return _p
    except Exception:
        return None

def _get_page_labels_via_pypdf(pdf_path: Path) -> Optional[List[str]]:
    """
    Use pypdf/PyPDF2 (if available) to read logical page labels, if any.
    Returns a list of labels (strings) or None if no labels or library not present.
    """
    mod = _import_pypdf()
    if not mod:
        return None
    try:
        with open(pdf_path, "rb") as fh:
            reader = mod.PdfReader(fh)
            # pypdf exposes .page_labels; older PyPDF2 may not.
            labels = None
            try:
                labels = reader.page_labels  # pypdf
            except Exception:
                # Fallback for older libs: inspect the catalog if accessible
                try:
                    root = reader.trailer.get("/Root", {})
                    if "/PageLabels" in root:
                        # Best effort: not all versions expose a friendly API here
                        # If present, we just return a placeholder list so caller knows labels exist.
                        labels = ["?"] * len(reader.pages)
                except Exception:
                    labels = None
            if labels:
                # Normalize to list of strings sized to page count
                if isinstance(labels, dict):
                    # pypdf sometimes returns a dict-like mapping; convert to ordered list
                    out = []
                    for i in range(len(reader.pages)):
                        out.append(str(labels.get(i, i+1)))
                    return out
                if isinstance(labels, list):
                    return [str(x) for x in labels]
                # Any truthy value implies there are labels; return a generic list
                return ["?"] * len(reader.pages)
            return None
    except Exception:
        return None

# --- main detector ---------------------------------------------------------

def detect_missing_pages(pdf: Path) -> Tuple[bool, str]:
    """
    Robust page-number detector (strengthened):
      • Prefer numbers from header/footer spans; whole-page or OCR-only hits
        are considered weak evidence.
      • If most pages lack header/footer numbers and no catalog PageLabels exist,
        flag as 'No page numbers printed'.
      • Flag odd/even-only sequences, gaps, suspicious jumps, poor monotonicity,
        and low uniqueness (repeating same numbers like a year).
    """
    re_page   = re.compile(r"\bpage\s+(\d+)\b(?:\s*(?:/|of)\s*(\d+))?", re.I)
    re_of     = re.compile(r"\b(\d+)\s+of\s+(\d+)\b", re.I)
    re_slash  = re.compile(r"\b(\d+)\s*/\s*(\d+)\b")
    re_brkt   = re.compile(r"^\s*[\(\[\{]?\s*(\d{1,4})\s*[\)\]\}]?\s*$")
    re_dash   = re.compile(r"^\s*[-–—]?\s*(\d{1,4})\s*[-–—]?\s*$")
    re_roman  = re.compile(r"^\s*(?=[IVXLCDMivxlcdm]{1,7}\s*$)[IVXLCDMivxlcdm]+\s*$")

    def extract_candidates(page) -> list[tuple[int, float]]:
        """
        Return [(number, provenance_bias)], lower bias is better.
        Provenance bias:
          0.0 = span located in header/footer zones (strong)
          0.3 = line from header/footer zones (good)
          0.6 = whole-page text (weak)
          1.0 = OCR strip only (weakest)
        """
        cands: list[tuple[int, float]] = []

        try:
            h = page.rect.height
            top_band = h * 0.20
            bot_band = h * 0.20
            d = page.get_text("dict")
            for b in d.get("blocks", []):
                for l in b.get("lines", []):
                    for s in l.get("spans", []):
                        y0 = s.get("bbox", [0,0,0,0])[1]
                        text = (s.get("text") or "").strip()
                        if not text:
                            continue
                        in_hf = (y0 <= top_band) or (y0 >= (h - bot_band))
                        if not in_hf:
                            continue
                        nums: list[int] = []
                        m = re_page.search(text) or re_of.search(text) or re_slash.search(text)
                        if m:
                            try:
                                nums.append(int(m.group(1)))
                            except Exception:
                                pass
                        else:
                            for token in re.findall(r"[IVXLCDM]+|\d{1,4}", text, flags=re.I):
                                if token.isdigit():
                                    nums.append(int(token))
                                else:
                                    rv = _roman_to_int(token)
                                    if rv is not None:
                                        nums.append(rv)
                        for n in nums:
                            cands.append((n, 0.0))
        except Exception:
            pass

        if not cands:
            try:
                h = page.rect.height
                top_band = h * 0.20
                bot_band = h * 0.20
                blocks = page.get_text("blocks")
                hf_lines = []
                for b in blocks:
                    x0,y0,x1,y1,txt = b[0],b[1],b[2],b[3],b[4]
                    if y0 <= top_band or y1 >= (h - bot_band):
                        hf_lines.extend((txt or "").splitlines())
                for line in hf_lines:
                    s = line.strip()
                    for token in re.findall(r"[IVXLCDM]+|\d{1,4}|page\s+\d+|\d+\s+of\s+\d+|\d+/\d+", s, flags=re.I):
                        if token.isdigit():
                            cands.append((int(token), 0.3))
                        elif token.lower().startswith("page"):
                            m = re_page.search(token)
                            if m: cands.append((int(m.group(1)), 0.3))
                        elif " of " in token:
                            m = re_of.search(token)
                            if m: cands.append((int(m.group(1)), 0.3))
                        elif "/" in token:
                            m = re_slash.search(token)
                            if m: cands.append((int(m.group(1)), 0.3))
                        else:
                            rv = _roman_to_int(token)
                            if rv is not None:
                                cands.append((rv, 0.3))
            except Exception:
                pass

        if not cands:
            for line in page.get_text("text").splitlines():
                s = line.strip()
                m = (re_page.search(s) or re_of.search(s) or re_slash.search(s) or
                     re_brkt.match(s) or re_dash.match(s))
                if m:
                    for g in m.groups() or ():
                        if g and g.isdigit():
                            cands.append((int(g), 0.6)); break
                    else:
                        for dgt in re.findall(r"\d{1,4}", s):
                            cands.append((int(dgt), 0.6))
                elif re_roman.match(s):
                    rv = _roman_to_int(s)
                    if rv is not None:
                        cands.append((rv, 0.6))

        if not cands:
            for where in ("bottom", "top"):
                o = _ocr_strip_numbers(page, band=0.20 if where=="bottom" else 0.16, where=where)
                cands.extend((n, 1.0) for n in o)

        best: dict[int, float] = {}
        for n, bias in cands:
            best[n] = min(best.get(n, 9.9), bias)
        return sorted(best.items(), key=lambda t: (t[1], t[0]))

    labels: list[Optional[int]] = []
    provenance: list[Optional[float]] = []  # track where the chosen label came from
    with fitz.open(pdf) as doc:
        for i in range(doc.page_count):
            expected = (labels[-1] + 1) if (labels and labels[-1] is not None) else (i + 1)
            cands = extract_candidates(doc[i])
            if cands:
                scored = [(abs(n - expected) + bias, n, bias) for n, bias in cands]
                scored.sort()
                label, bias = scored[0][1], scored[0][2]
                if len(scored) >= 2:
                    (s0, n0, b0), (s1, n1, b1) = scored[0], scored[1]
                    if abs(n1 - expected) in (0, 1) and s1 <= s0 + 0.35:
                        label, bias = n1, b1
                labels.append(label)
                provenance.append(bias)
            else:
                labels.append(None)
                provenance.append(None)

    # If absolutely nothing was detected at all:
    present = [n for n in labels if n is not None]
    if not present:
        # consult catalog page labels if any
        cat_labels = _get_page_labels_via_pypdf(pdf)
        if not cat_labels:
            return (True, "No page numbers detected in content and no PDF catalog page labels present")
        else:
            # Having only catalog labels but nothing printed is still an issue (no printed numbers)
            return (True, "No printed page numbers detected; only PDF catalog page labels present")

    bits: list[str] = []
    N = len(labels)
    # Header/footer strength
    hf_hits = sum(1 for b in provenance if (b is not None and b <= 0.3))
    any_hits = sum(1 for b in provenance if (b is not None))
    hf_ratio = hf_hits / max(1, N)
    any_ratio = any_hits / max(1, N)

    # If most pages don't have strong header/footer numbers, treat as 'no printed page numbers'
    # unless the catalog provides explicit page labels (which we still consider "not printed").
    cat_labels = _get_page_labels_via_pypdf(pdf)
    if hf_ratio < 0.60:
        if not cat_labels:
            bits.append("Printed page numbers not found on a sufficient number of pages (header/footer evidence < 60%)")
        else:
            bits.append("No printed page numbers; only PDF catalog page labels present")

    # Odd/even-only quick checks
    odd_only  = (len(present) >= 4 and all(n % 2 == 1 for n in present))
    even_only = (len(present) >= 4 and all(n % 2 == 0 for n in present))
    if odd_only:  bits.append("Odd pages only")
    if even_only: bits.append("Even pages only")

    # Missing printed numbers on some pages
    if any_ratio < 0.95:  # tolerate small OCR/text misses
        missing_cnt = N - any_hits
        bits.append(f"Printed number missing on {missing_cnt}/{N} pages")

    # Suspicious jumps and gaps
    if not (odd_only or even_only):
        prev = None
        jumps = []
        for idx, n in enumerate(labels, start=1):
            if n is None:
                prev = None
                continue
            if prev is not None:
                delta = n - prev
                if abs(delta) > 2:  # tighten threshold; +1 expected
                    jumps.append(f"{prev}→{n} (near physical page {idx})")
            prev = n
        if jumps:
            bits.append(f"Suspicious numbering jumps: {', '.join(jumps[:6])}{'…' if len(jumps) > 6 else ''}")

        s = sorted(set(present))
        gaps = [n for n in range(s[0], s[-1]+1) if n not in s]
        if gaps:
            bits.append(f"Page numbering gaps: {gaps[:10]}{'…' if len(gaps) > 10 else ''}")

    # Monotonicity & uniqueness quality
    # Measure fraction of +1 transitions among all adjacent pairs where both sides had labels
    pairs = []
    prev_n = None
    for n in labels:
        if n is not None and prev_n is not None:
            pairs.append(1 if (n - prev_n) == 1 else 0)
        if n is not None:
            prev_n = n
        else:
            prev_n = None
    if pairs:
        plus1_ratio = sum(pairs) / len(pairs)
        if plus1_ratio < 0.50:
            bits.append(f"Poor sequential consistency (+1 transitions in only {int(plus1_ratio*100)}% of adjacent pages)")

    # If the set of numbers is too small relative to page count, they may be repeated non-page numbers (e.g., a year)
    uniq_ratio = len(set(present)) / max(1, any_hits)
    if uniq_ratio < 0.50:
        bits.append("Low uniqueness of detected numbers (likely non-page numbers repeating, e.g., a year or figure code)")

    # Final decision
    bits = sorted(set(bits))
    has_issue = bool(bits)
    reason = "; ".join(bits) if has_issue else ""
    return (has_issue, reason or "No issues detected")

def find_section_presence(text: str, compiled_syn: Dict[str, List[re.Pattern]]) -> Tuple[Dict[str, bool], Dict[str, List[str]]]:
    """Return (found_map, which_map). We record full matches (group(0)) for safety."""
    found_map: Dict[str, bool] = {}
    which_map: Dict[str, List[str]] = {}
    for section, regs in compiled_syn.items():
        labels = set()
        hit = False
        for rg in regs:
            for m in rg.finditer(text):
                s = m.group(0)
                if s:
                    labels.add(s.strip())
                    hit = True
        found_map[section] = hit
        which_map[section] = sorted(labels)
    return found_map, which_map

# ===============================
# Issues loader (sheet required), with optional defaults merge
# ===============================
def load_issues_xlsx(path: Optional[Path]) -> pd.DataFrame:
    if not path:
        raise ValueError("Issues.xlsx is required. Please pass --issues-xlsx <file>")
    if not path.exists():
        raise FileNotFoundError(f"Issues file not found: {path}")
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]
    if df.empty:
        raise ValueError("Issues.xlsx is empty. Please add rules.")
    return df

def build_issue_rules_from_sheet(df: pd.DataFrame) -> List[dict]:
    rules: List[dict] = []
    for idx, row in df.iterrows():
        rule: Dict[str, object] = {}
        # IDs/names
        rule["IssueID"] = (str(row.get("IssueID")).strip()
                           if pd.notna(row.get("IssueID")) else f"ISSUE_{idx+1}")
        iname = row.get("IssueName")
        rule["IssueName"] = (str(iname).strip() if pd.notna(iname) else "") or rule["IssueID"]

        # Type
        rt = row.get("RuleType")
        if pd.isna(rt):
            raise ValueError(f"Row {idx+2}: RuleType is required (keyword|missing_section|missing_pages).")
        rt = str(rt).strip().lower()
        if rt not in {"keyword", "missing_section", "missing_pages"}:
            raise ValueError(f"Row {idx+2}: invalid RuleType '{rt}'.")
        rule["RuleType"] = rt

        # Scope
        scope = row.get("Scope")
        rule["Scope"] = str(scope).strip().lower() if pd.notna(scope) else "any"

        # MustBePresent
        mbp = row.get("MustBePresent")
        rule["MustBePresent"] = bool(mbp) if pd.notna(mbp) else True

        # Keyword regex
        regex = row.get("Regex", "")
        if pd.notna(regex):
            regex = str(regex).strip()
        else:
            regex = ""
        rule["Regex"] = regex

        # Section (for missing_section)
        sec = row.get("Section", "")
        if pd.notna(sec):
            sec = str(sec).strip()
        else:
            sec = ""
        rule["Section"] = sec

        # Optional SectionRegex (for missing_section override)
        srgx = row.get("SectionRegex", "")
        if pd.notna(srgx):
            srgx = str(srgx).strip()
        else:
            srgx = ""
        rule["SectionRegex"] = srgx

        # Validation
        if rt == "keyword" and not regex:
            typer.echo(f"[WARN] Row {idx+2}: keyword rule without Regex will never match.")
        if rt == "missing_section" and not sec and not srgx:
            raise ValueError(f"Row {idx+2}: missing_section needs Section or SectionRegex.")

        rules.append(rule)
    return rules

def eval_keyword_rule(text: str, regex: str, must_be_present: bool,
                      scope: str, section_texts: Optional[Dict[str, str]] = None) -> Tuple[bool, List[str]]:
    flags = re.I | re.M
    pat = re.compile(regex, flags) if regex else None
    matched: List[str] = []

    haystack = text
    if scope.startswith("section:") and section_texts:
        sec = scope.split(":", 1)[1].strip()
        haystack = section_texts.get(sec, "")

    if pat:
        matches = [m.group(0) for m in pat.finditer(haystack)]
        if matches:
            matched = sorted(set(m.strip() for m in matches if m))

    if must_be_present:
        return (len(matched) > 0, matched)
    else:
        return (len(matched) == 0, matched)

# ===============================
# CLI (single command)
# ===============================
@app.command()
def main(
    input_dir: Path = typer.Option(..., "--input-dir", "-i", help="Folder containing PDFs (recursively scanned)"),
    out_sections_csv: Path = typer.Option("sections_presence.csv", "--out-sections-csv", help="Output CSV for section presence"),
    out_issues_csv: Path = typer.Option("issues_scan.csv", "--out-issues-csv", help="Output CSV for issue rules"),
    synonyms_csv: Optional[Path] = typer.Option(None, "--synonyms-csv", "-s", help="Optional synonyms CSV (Section,KeywordRegex)"),
    issues_xlsx: Optional[Path] = typer.Option(..., "--issues-xlsx", "-x", help="Issues Excel with rules (REQUIRED)"),
    ocr_cache_dir: Path = typer.Option("ocr_out", "--ocr-cache-dir", help="Where to put OCR'ed PDFs if ocrmypdf is available"),
    pages_limit: int = typer.Option(0, "--pages-limit", help="0 = all pages; otherwise limit pages per PDF for speed"),
    check_pages: bool = typer.Option(True, "--check-pages/--no-check-pages", help="Enable 'missing_pages' rules if present or defaulted"),
    with_defaults: bool = typer.Option(False, "--with-defaults/--no-defaults", help="Merge built-in default rules with Issues.xlsx"),
    summary_csv: Path = typer.Option("out/summary_report.csv", "--out-summary-csv", help="Output CSV for summary report"),
):
    # 1) Load Section Synonyms
    syn_map = load_synonyms_csv(synonyms_csv)
    compiled_syn = compile_section_synonyms(syn_map)
    section_key_map = build_section_key_map(list(compiled_syn.keys()))

    # 2) Load Issues.xlsx + optional defaults
    issues_df = load_issues_xlsx(issues_xlsx)
    rules = build_issue_rules_from_sheet(issues_df)
    if with_defaults:
        # merge defaults; keep sheet precedence for same IssueID
        seen = {r["IssueID"] for r in rules}
        for r in DEFAULT_ISSUE_RULES:
            if r["IssueID"] not in seen:
                rules.append(r)

    # 2.1) Normalize rule IDs: map any MUST_HAVE_* IDs to MISSING_* for consistency
    # We honour specific mappings first (e.g. MUST_HAVE_VALUE -> MISSING_VALUEOFSTUDY) then
    # generic MUST_HAVE_* -> MISSING_* conversions.  This ensures downstream code and
    # summaries operate on a single naming convention.
    rename_map = {
        "MUST_HAVE_TIMELINE": "MISSING_TIMELINE",
        "MUST_HAVE_BUDGET": "MISSING_BUDGET",
        "MUST_HAVE_VALUE": "MISSING_VALUEOFSTUDY",
        "MUST_HAVE_OBJECTIVES": "MISSING_OBJECTIVES",
    }
    for rule in rules:
        issue_id = rule.get("IssueID", "")
        # Apply explicit mapping if present
        if issue_id in rename_map:
            rule["IssueID"] = rename_map[issue_id]
        # Generic fallback: convert MUST_HAVE_* to MISSING_*
        elif isinstance(issue_id, str) and issue_id.startswith("MUST_HAVE_"):
            rule["IssueID"] = "MISSING_" + issue_id[len("MUST_HAVE_"):]

    # 3) Discover PDFs
    pdfs: List[Path] = []
    for root, _, files in os.walk(input_dir):
        for f in files:
            if f.lower().endswith(".pdf"):
                pdfs.append(Path(root) / f)
    pdfs.sort()

    sec_rows: List[dict] = []
    issue_rows: List[dict] = []
    # Dictionaries to accumulate missing sections and issue pass/fail statuses per PDF
    missing_sections_by_pdf: Dict[str, List[str]] = {}
    issues_status_by_pdf: Dict[str, Dict[str, str]] = {}
    # Dictionaries to accumulate aggregated matched keywords and pass statuses per PDF
    aggregated_keywords_by_pdf: Dict[str, Dict[str, str]] = {}
    aggregated_passyn_by_pdf: Dict[str, Dict[str, str]] = {}

    # 4) Process
    for idx, pdf in enumerate(pdfs, 1):
        typer.echo(f"[{idx}/{len(pdfs)}] Scanning {pdf.name} ...")
        try:
            prepared = ensure_ocr_pdf(pdf, Path(ocr_cache_dir))
            text = pdf_to_text(prepared, pages_limit if pages_limit > 0 else None)

            # Section presence
            found_map, which_map = find_section_presence(text, compiled_syn)
            per_section_text: Dict[str, str] = {sec: text for sec in compiled_syn.keys()}  # simple scope
            # Collect missing sections and initialise issue status map for this PDF
            missing_list: List[str] = [sec for sec in compiled_syn.keys() if not found_map.get(sec, False)]
            missing_sections_by_pdf[pdf.name] = missing_list
            # Prepare a dict to record issue pass/fail for this PDF
            issues_status_by_pdf[pdf.name] = {}
            # Initialize aggregated matched keywords and pass statuses dicts for this PDF
            aggregated_keywords_by_pdf[pdf.name] = {}
            aggregated_passyn_by_pdf[pdf.name] = {}

            # sections_presence.csv
            for sec in compiled_syn.keys():
                mk = which_map.get(sec, [])
                sec_rows.append({
                    "PDF_FILE_NAME": pdf.name,
                    "SECTION": sec,
                    "MATCHED_KEYWORDS": ", ".join(mk) if mk else "",
                    "PRESENT_YN": "YES" if found_map.get(sec, False) else "NO",
                    "REMARKS": "",
                })

            # pre-compute missing_pages if such rule exists
            mp_has_issue, mp_rem = (False, "")
            if check_pages and any(r["RuleType"] == "missing_pages" for r in rules):
                mp_has_issue, mp_rem = detect_missing_pages(prepared)

            for rule in rules:
                rtype = str(rule["RuleType"])
                must_present = bool(rule.get("MustBePresent", True))
                scope = str(rule.get("Scope", "any")).lower()
                issue_id = rule.get("IssueID", "")
                issue_name = rule.get("IssueName", issue_id)
                matched_labels: List[str] = []
                pass_yn = "YES"
                remarks = ""
                if rtype == "missing_pages":
                    if not check_pages:
                        pass_yn = "N/A"
                    else:
                        # mp_has_issue=True 表示有缺頁碼、只有單數頁或偶數頁等問題
                        # 因此 pass_yn 應為 "NO"
                        pass_yn = "NO" if mp_has_issue else "YES"
                    remarks = mp_rem
                    # Record pass/fail status for this PDF and issue
                    issues_status_by_pdf[pdf.name][issue_id] = pass_yn
                elif rtype == "missing_section":
                    sec_raw = str(rule.get("Section", "")).strip()
                    sec_regex = str(rule.get("SectionRegex", "")).strip()
                    present = False
                    if sec_regex:
                        try:
                            pat = re.compile(sec_regex, re.I | re.M)
                            hits = [m.group(0).strip() for m in pat.finditer(text)]
                            if hits:
                                present = True
                                matched_labels = sorted(set(hits))
                        except re.error as e:
                            typer.echo(f"[WARN] Invalid SectionRegex for '{sec_raw}': {sec_regex} ({e}); falling back to synonyms")
                            sec_regex = ""
                    if not sec_regex:
                        if not sec_raw:
                            typer.echo("[WARN] missing_section rule without 'Section' — skipped.")
                            continue
                        canon = section_key_map.get(_norm_key(sec_raw), sec_raw)
                        present = bool(found_map.get(canon, False))
                        matched_labels = which_map.get(canon, []) if present else []
                    pass_yn = "YES" if (present == must_present) else "NO"
                    if not present:
                        remarks = "Section not detected"
                    # Record pass/fail status for this PDF and issue
                    issues_status_by_pdf[pdf.name][issue_id] = pass_yn
                else:
                    ok, labels = eval_keyword_rule(text, str(rule.get("Regex", "")), must_present, scope, per_section_text)
                    pass_yn = "YES" if ok else "NO"
                    matched_labels = labels
                    # Record pass/fail status for this PDF and issue
                    issues_status_by_pdf[pdf.name][issue_id] = pass_yn

                # Append aggregated matched keywords and pass status for this rule, keyed by issue ID
                kw_string = ", ".join(matched_labels) if matched_labels else ""
                aggregated_keywords_by_pdf[pdf.name][issue_id] = kw_string
                aggregated_passyn_by_pdf[pdf.name][issue_id] = pass_yn
                issue_rows.append({
                    "PDF_FILE_NAME": pdf.name,
                    "ISSUE_ID": issue_id,
                    "ISSUE_NAME": issue_name,
                    "RULE_TYPE": rtype,
                    "RULE_SCOPE": scope,
                    "MUST_BE_PRESENT": "YES" if must_present else "NO",
                    "MATCHED_KEYWORDS": ", ".join(matched_labels) if matched_labels else "",
                    "PASS_YN": pass_yn,
                    "REMARKS": remarks,
                })
        except Exception as e:
            for sec in compiled_syn.keys():
                sec_rows.append({
                    "PDF_FILE_NAME": pdf.name,
                    "SECTION": sec,
                    "MATCHED_KEYWORDS": "",
                    "PRESENT_YN": "ERROR",
                    "REMARKS": f"{type(e).__name__}: {e}",
                })
            issue_rows.append({
                "PDF_FILE_NAME": pdf.name,
                "ISSUE_ID": "RUNTIME_ERROR",
                "ISSUE_NAME": "Script error while processing file",
                "RULE_TYPE": "runtime",
                "RULE_SCOPE": "any",
                "MUST_BE_PRESENT": "",
                "MATCHED_KEYWORDS": "",
                "PASS_YN": "ERROR",
                "REMARKS": f"{type(e).__name__}: {e}",
            })

    # 5) Write outputs
    pd.DataFrame(sec_rows).to_csv(out_sections_csv, index=False)
    pd.DataFrame(issue_rows).to_csv(out_issues_csv, index=False)
    typer.secho(f"Wrote {out_sections_csv} and {out_issues_csv} (files scanned: {len(pdfs)})", fg=typer.colors.GREEN)

    # 6) Build a summary combining missing sections and issue failures per PDF
    summary_rows: List[Dict[str, str]] = []
    # Order of issue IDs to append in remarks.  These IDs correspond to the
    # normalized names (MISSING_*) rather than any legacy MUST_HAVE_* names.
    summary_issue_order = [
        "MISSING_TIMELINE",
        "MISSING_BUDGET",
        "MISSING_VALUEOFSTUDY",
        "MISSING_OBJECTIVES",
        "MISSING_PAGES",
    ]
    for pdf in pdfs:
        fname = pdf.name
        remarks_parts: List[str] = []
        # Missing sections first
        missing_secs = missing_sections_by_pdf.get(fname, [])
        if missing_secs:
            # Prefix missing section reports with English label: MISSING SECTIONS
            remarks_parts.append("MISSING SECTIONS: " + ", ".join(missing_secs))
        # Then failed issues in specified order
        statuses = issues_status_by_pdf.get(fname, {})
        for issue_id in summary_issue_order:
            status = statuses.get(issue_id)
            if status in {"NO", "ERROR"}:
                remarks_parts.append(issue_id)
        # Aggregate matched keywords and pass statuses for this PDF.
        # Build a list like "ISSUE_ID: [keywords]" for each issue.
        kw_entries: List[str] = []
        kw_dict = aggregated_keywords_by_pdf.get(fname, {}) or {}
        # Only include entries with matched keywords; remove prefixes like MISS_, MISSING_, MISS_SEC_ from issue IDs
        for _issue_id, _kw_string in kw_dict.items():
            # Skip if no keywords matched
            if not _kw_string or not _kw_string.strip():
                continue
            clean_id = _issue_id
            # Remove specific prefixes
            if clean_id.startswith("MISS_SEC_"):
                clean_id = clean_id[len("MISS_SEC_"):]
            elif clean_id.startswith("MISS_"):
                clean_id = clean_id[len("MISS_"):]
            elif clean_id.startswith("MISSING_"):
                clean_id = clean_id[len("MISSING_"):]
            kw_entries.append(f"[{clean_id}]:{_kw_string}")
        aggregated_kw_string = "; ".join(kw_entries)
        # Determine aggregated PASS_YN: YES only if all statuses are YES or N/A; otherwise NO.
        passes_dict = aggregated_passyn_by_pdf.get(fname, {}) or {}
        if passes_dict:
            aggregated_pass_string = "YES" if all(_status in {"YES", "N/A"} for _status in passes_dict.values()) else "NO"
        else:
            aggregated_pass_string = "YES"
        summary_rows.append({
            "PDF_FILE_NAME": fname,
            "MATCHED_KEYWORDS": aggregated_kw_string,
            "ALL_PASS": aggregated_pass_string,
            "REMARKS": ", ".join(remarks_parts)
        })
    # Write summary report to the requested CSV path.  Use the user-supplied
    # --out-summary-csv argument rather than a hard-coded location.
    pd.DataFrame(summary_rows).to_csv(summary_csv, index=False)
    typer.secho(f"Wrote {summary_csv} (files scanned: {len(pdfs)})", fg=typer.colors.GREEN)


if __name__ == "__main__":
    app()
