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
    {"IssueID": "MUST_HAVE_TIMELINE", "IssueName": "Timeline/schedule not found",
     "RuleType": "keyword",
     "Regex": r"\b(timeline|schedule|gantt|milestone|work\s*plan|workplan|time\s*line)\b",
     "MustBePresent": True, "Scope": "any"},
    {"IssueID": "MUST_HAVE_BUDGET", "IssueName": "Budget/cost not found",
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

def detect_missing_pages(pdf: Path) -> Tuple[bool, str]:
    """
    Return (has_issue, remarks) based on printed page numbers.

    Improvements:
    - Wider patterns (brackets, en/em dashes, 3/10, 3 of 10, roman numerals).
    - If header/footer yields few candidates, also search full page body.
    - Greedy selection prefers the candidate closest to expected (prev+1),
      falling back to physical page index (i+1).
    - Flags suspicious jumps (e.g., … 10 -> 57) as a remark.
    """
    # Patterns: "page 3", "3 of 10", "3/10", "(3)", "[3]", "— 3 —", lone "3", roman numerals
    re_page   = re.compile(r"\bpage\s+(\d+)\b(?:\s*(?:/|of)\s*(\d+))?", re.I)
    re_of     = re.compile(r"\b(\d+)\s+of\s+(\d+)\b", re.I)
    re_slash  = re.compile(r"\b(\d+)\s*/\s*(\d+)\b")
    re_brkt   = re.compile(r"^\s*[\(\[\{]?\s*(\d{1,4})\s*[\)\]\}]?\s*$")
    re_dash   = re.compile(r"^\s*[-–—]?\s*(\d{1,4})\s*[-–—]?\s*$")
    re_roman  = re.compile(r"^\s*(?=[IVXLCDMivxlcdm]{1,7}\s*$)[IVXLCDMivxlcdm]+\s*$")

    def roman_to_int(s: str) -> Optional[int]:
        m = {'I':1,'V':5,'X':10,'L':50,'C':100,'D':500,'M':1000}
        s = s.upper().strip()
        if not s or not re_roman.match(s): return None
        val = 0; prev = 0
        for ch in s:
            cur = m.get(ch, 0)
            if cur > prev: val += cur - 2*prev
            else: val += cur
            prev = cur
        return val if 1 <= val <= 4000 else None

    labels: List[Optional[int]] = []

    with fitz.open(pdf) as doc:
        for i in range(doc.page_count):
            page = doc[i]
            # 1) Prefer header/footer lines; if scarce, include more lines
            try:
                h = page.rect.height
                top_band = h * 0.25
                bot_band = h * 0.25
                blocks = page.get_text("blocks")
                hf_lines = []
                for b in blocks:
                    x0,y0,x1,y1,txt = b[0],b[1],b[2],b[3],b[4]
                    if y0 <= top_band or y1 >= (h - bot_band):
                        hf_lines.extend(txt.splitlines())
                all_lines = page.get_text("text").splitlines()
                # If header/footer gave too few lines, widen search to whole page
                lines = hf_lines if len(hf_lines) >= 4 else all_lines
            except Exception:
                lines = page.get_text("text").splitlines()

            # 2) Collect candidates on this page
            cands: List[int] = []
            for line in lines:
                s = line.strip()
                m = re_page.search(s)
                if m:
                    cands.append(int(m.group(1))); continue
                m = re_of.search(s)
                if m:
                    cands.append(int(m.group(1))); continue
                m = re_slash.search(s)
                if m:
                    cands.append(int(m.group(1))); continue
                m = re_brkt.match(s)
                if m:
                    cands.append(int(m.group(1))); continue
                m = re_dash.match(s)
                if m:
                    cands.append(int(m.group(1))); continue
                if re_roman.match(s):
                    rv = roman_to_int(s)
                    if rv is not None:
                        cands.append(rv); continue

            # 3) Choose the best label for this page
            # Prefer the candidate closest to expected (prev+1); else use physical (i+1)
            if cands:
                if labels and labels[-1] is not None:
                    expected = labels[-1] + 1
                else:
                    expected = i + 1
                label = min(cands, key=lambda x: abs(x - expected))
                labels.append(label)
            else:
                labels.append(None)

    present = [n for n in labels if n is not None]

    # --- no numbers at all
    if not present:
        return (True, "No page numbers detected")

    bits: List[str] = []

    # odd/even-only
    odd_only  = present and all(n % 2 == 1 for n in present)
    even_only = present and all(n % 2 == 0 for n in present)
    if odd_only:
        bits.append("Odd pages only")
    if even_only:
        bits.append("Even pages only")

    # pages that lack a printed number
    if len(present) < len(labels):
        bits.append(f"Printed number missing on {len(labels)-len(present)}/{len(labels)} pages")

    # suspicious jumps (difference > 3 between consecutive present labels, ignoring odd/even-only docs)
    if not (odd_only or even_only):
        prev = None
        for idx, n in enumerate(labels, start=1):
            if n is None:
                continue
            if prev is not None:
                delta = n - prev
                if abs(delta) > 3:
                    bits.append(f"Suspicious numbering jump near page {idx}: {prev} → {n}")
                    # don’t spam; one is enough to flag, but keep going for completeness
            prev = n

    # gaps in a contiguous sequence (skip on odd/even-only)
    if not (odd_only or even_only):
        s = sorted(set(present))
        if s:
            gaps = [n for n in range(s[0], s[-1] + 1) if n not in s]
            if gaps:
                bits.append(f"Page numbering gaps: {gaps[:10]}{'…' if len(gaps) > 10 else ''}")

    return (bool(bits), "; ".join(sorted(set(bits))))

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

    # 3) Discover PDFs
    pdfs: List[Path] = []
    for root, _, files in os.walk(input_dir):
        for f in files:
            if f.lower().endswith(".pdf"):
                pdfs.append(Path(root) / f)
    pdfs.sort()

    sec_rows: List[dict] = []
    issue_rows: List[dict] = []

    # 4) Process
    for idx, pdf in enumerate(pdfs, 1):
        typer.echo(f"[{idx}/{len(pdfs)}] Scanning {pdf.name} ...")
        try:
            prepared = ensure_ocr_pdf(pdf, Path(ocr_cache_dir))
            text = pdf_to_text(prepared, pages_limit if pages_limit > 0 else None)

            # Section presence
            found_map, which_map = find_section_presence(text, compiled_syn)
            per_section_text: Dict[str, str] = {sec: text for sec in compiled_syn.keys()}  # simple scope

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
                        remarks = "Page check disabled"
                    else:
                        pass_yn = "NO" if mp_has_issue else "YES" if must_present else ("YES" if mp_has_issue else "NO")
                        remarks = mp_rem or remarks
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
                else:
                    ok, labels = eval_keyword_rule(text, str(rule.get("Regex", "")), must_present, scope, per_section_text)
                    pass_yn = "YES" if ok else "NO"
                    matched_labels = labels

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


if __name__ == "__main__":
    app()
