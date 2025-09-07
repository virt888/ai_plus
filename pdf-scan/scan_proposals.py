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

def _norm_key(s: str) -> str:
    """Lowercase + remove non-alphanumerics, e.g. 'Outcomes & Value' -> 'outcomesvalue'."""
    import re
    return re.sub(r"[^a-z0-9]+", "", s.lower())

def build_section_key_map(section_names):
    """
    Build a map from normalized aliases -> canonical section key used by synonyms.
    Also includes helpful aliases that should resolve to 'OutcomesValue'.
    """
    m = {_norm_key(k): k for k in section_names}
    # Common human variants all map to OutcomesValue
    for alias in [
        "outcomesvalue", "outcomesandvalue", "outcome", "expectedoutcomes",
        "keyoutcomes", "outcomesimpact", "results", "deliverables"
    ]:
        m.setdefault(alias, "OutcomesValue")
    return m

app = typer.Typer(add_completion=False)

# ===============================
# Global Config (edit if needed)
# ===============================

# OCR (Tesseract) config
TESS_LANG = "eng"          # e.g., "chi_tra+eng" or "chi_sim+eng" if Chinese headings appear
TESS_CFG  = r"--oem 3 --psm 6"
DPI       = 150            # page raster DPI for OCR

# Default Section Synonyms (used for section presence + "missing_section" rules)
# NOTE: OutcomesValue pattern is expanded to catch "Expected Outcomes", numbered headings, etc.
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
    "OutcomesValue": [
        r"^\s*(\d+[\.\)]\s*)?(project\s+)?(expected\s+|anticipated\s+)?(outcomes?|results?|deliverables|outputs?|impact|findings|analysis)\b",
    ],
    "References": [
        r"^\s*(references|bibliography|works cited)\b",
    ],
}

# Built-in default issue rules (merged with Issues.xlsx if provided)
# - missing_pages: detect odd/even-only or numbering gaps
# - missing_section: for all default sections
# - keyword presence: timeline/schedule and budget/cost
DEFAULT_ISSUE_RULES = [
    # Missing pages heuristics
    {"IssueID": "MISSING_PAGES", "IssueName": "Missing / odd-even-only pages",
     "RuleType": "missing_pages", "Regex": "", "MustBePresent": True, "Scope": "any"},

    # Missing section presence
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

    # Keyword must-have (typical FHDC flags)
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

def compile_section_synonyms(mapping: Dict[str, List[str]]) -> Dict[str, List[re.Pattern]]:
    return {k: [re.compile(p, re.I | re.M) for p in v] for k, v in mapping.items()}

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
    """
    If the PDF has no text layer and `ocrmypdf` exists, produce an OCR'ed PDF in cache dir.
    Otherwise return original (we will OCR per page on-the-fly).
    """
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
    return pdf  # fallback

def page_to_text_via_ocr(page) -> str:
    mat = fitz.Matrix(DPI/72, DPI/72)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    return pytesseract.image_to_string(img, lang=TESS_LANG, config=TESS_CFG)

def pdf_to_text(pdf: Path, pages_limit: Optional[int]) -> str:
    """
    Extract text; OCR pages with no text.
    """
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
    Heuristic: detect printed page numbering gaps and odd/even-only sequences.
    Returns (has_issue, remarks).
    """
    re_num = re.compile(r"\bpage\s+(\d+)\b(?:\s*of\s*(\d+))?", re.I)
    re_lone = re.compile(r"^\s*(\d{1,3})\s*$")
    nums: List[Optional[int]] = []
    with fitz.open(pdf) as doc:
        for i in range(doc.page_count):
            text = doc[i].get_text("text")
            cands = []
            for line in text.splitlines():
                m = re_num.search(line)
                if m:
                    cands.append(int(m.group(1)))
                    continue
                m2 = re_lone.match(line)
                if m2:
                    cands.append(int(m2.group(1)))
            nums.append(min(cands, key=lambda x: abs(x - (i + 1))) if cands else None)

    present = [n for n in nums if n is not None]
    if not present:
        return (False, "")  # inconclusive
    s = sorted(set(present))
    gaps = [n for n in range(s[0], s[-1] + 1) if n not in s]
    odd_only = present and all(n % 2 == 1 for n in present)
    even_only = present and all(n % 2 == 0 for n in present)

    bits = []
    if gaps:
        bits.append(f"Page numbering gaps: {gaps[:10]}{'…' if len(gaps) > 10 else ''}")
    if odd_only or even_only:
        bits.append("Odd pages only" if odd_only else "Even pages only")
    return (bool(bits), "; ".join(bits))

def find_section_presence(text: str, compiled_syn: Dict[str, List[re.Pattern]]) -> Tuple[Dict[str, bool], Dict[str, List[str]]]:
    found_map: Dict[str, bool] = {}
    which_map: Dict[str, List[str]] = {}
    for section, regs in compiled_syn.items():
        labels = set()
        hit = False
        for rg in regs:
            for m in rg.finditer(text):
                s = m.group(0)  # <-- always safe
                if s:
                    labels.add(s.strip())
                    hit = True
        found_map[section] = hit
        which_map[section] = sorted(labels)
    return found_map, which_map

# ===============================
# Issues Loader & Evaluators
# ===============================

def load_synonyms_csv(path: Optional[Path]) -> Dict[str, List[str]]:
    """Return {Section: [regex,...]} with robust handling of blanks and bad rows."""
    if not path or not path.exists():
        return {k: v[:] for k, v in DEFAULT_SECTION_SYNONYMS.items()}

    sdf = pd.read_csv(path)
    sdf = sdf.rename(columns=lambda c: str(c).strip())
    if "Section" not in sdf.columns or "KeywordRegex" not in sdf.columns:
        raise ValueError("synonyms.csv must have columns: Section, KeywordRegex")

    # Normalize and drop blanks
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

def load_issues_xlsx(path: Optional[Path]) -> pd.DataFrame:
    if path and path.exists():
        try:
            df = pd.read_excel(path)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception as e:
            typer.echo(f"[WARN] Failed to read Issues.xlsx: {e}")
    return pd.DataFrame()

def build_issue_rules(df: pd.DataFrame, default_sections: List[str]) -> List[dict]:
    rules: List[dict] = []
    has_missing_pages = False
    has_missing_section = False
    has_keyword = False

    if not df.empty:
        for idx, row in df.iterrows():
            rule: Dict[str, object] = {}
            rule["IssueID"] = (str(row.get("IssueID")).strip()
                               if pd.notna(row.get("IssueID")) else f"ISSUE_{idx+1}")
            iname = row.get("IssueName")
            rule["IssueName"] = (str(iname).strip() if pd.notna(iname) else "") or rule["IssueID"]

            rt = row.get("RuleType")
            rt = str(rt).strip().lower() if pd.notna(rt) else "keyword"
            rule["RuleType"] = rt

            regex = row.get("Regex", "")
            if pd.isna(regex):
                regex = ""
            kw2 = row.get("Keyword", "")
            if pd.notna(kw2) and not regex:
                regex = str(kw2)
            rule["Regex"] = str(regex).strip()

            mbp = row.get("MustBePresent")
            rule["MustBePresent"] = bool(mbp) if pd.notna(mbp) else True

            scope = row.get("Scope")
            rule["Scope"] = str(scope).strip().lower() if pd.notna(scope) else "any"

            if rt == "missing_pages":
                has_missing_pages = True
            if rt == "missing_section":
                has_missing_section = True
                sec = row.get("Section")
                sec = str(sec).strip() if pd.notna(sec) else ""
                if not sec:
                    typer.echo("[WARN] missing_section rule without 'Section' – skipped.")
                    continue
                rule["Section"] = sec
            if rt == "keyword":
                has_keyword = True

            rules.append(rule)

    # Add built-in defaults if not overridden by sheet
    if not has_missing_pages:
        rules.append([r for r in DEFAULT_ISSUE_RULES if r["RuleType"] == "missing_pages"][0])

    if not has_missing_section:
        for sec in default_sections:
            rules.append({
                "IssueID": f"MISS_SEC_{sec}",
                "IssueName": f"Missing {sec} section",
                "RuleType": "missing_section",
                "Section": sec,
                "Regex": "",
                "MustBePresent": True,
                "Scope": "any",
            })

    if not has_keyword or df.empty:
        for r in DEFAULT_ISSUE_RULES:
            if r["RuleType"] == "keyword":
                rules.append(r)

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
            matched = sorted(set(matches))

    if must_be_present:
        return (len(matched) > 0, matched)
    else:
        return (len(matched) == 0, matched)

# ===============================
# CLI (single command — no subcommands)
# ===============================

@app.command()
def main(
    input_dir: Path = typer.Option(..., "--input-dir", "-i", help="Folder containing PDFs (recursively scanned)"),
    out_sections_csv: Path = typer.Option("sections_presence.csv", "--out-sections-csv", help="Output CSV for section presence"),
    out_issues_csv: Path = typer.Option("proposal_scan_result.csv", "--out-issues-csv", help="Output CSV for issue rules"),
    synonyms_csv: Optional[Path] = typer.Option(None, "--synonyms-csv", "-s", help="Optional synonyms CSV (Section,KeywordRegex)"),
    issues_xlsx: Optional[Path] = typer.Option(None, "--issues-xlsx", "-x", help="Issues Excel to drive rule checks"),
    ocr_cache_dir: Path = typer.Option("ocr_out", "--ocr-cache-dir", help="Where to put OCR'ed PDFs if ocrmypdf is available"),
    pages_limit: int = typer.Option(0, "--pages-limit", help="0 = all pages; otherwise limit pages per PDF for speed"),
    check_pages: bool = typer.Option(True, "--check-pages/--no-check-pages", help="Heuristically flag missing/odd-even pages"),
):
    # 1) Load Section Synonyms (CSV or defaults)
    try:
        syn_map = load_synonyms_csv(synonyms_csv)
        compiled_syn = compile_section_synonyms(syn_map)
        section_key_map = build_section_key_map(list(compiled_syn.keys()))
    except Exception as e:
        typer.echo(f"[ERROR] Failed to load synonyms: {e}")
        raise

    # 2) Load Issues (XLSX) and merge with defaults
    issues_df = load_issues_xlsx(issues_xlsx)
    default_sections = list(compiled_syn.keys())
    issue_rules = build_issue_rules(issues_df, default_sections)

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

            # (Optional) per-section texts (currently naive — whole text; can segment later)
            per_section_text: Dict[str, str] = {sec: text for sec in compiled_syn.keys()}

            # Rows for sections_presence.csv
            for sec in compiled_syn.keys():
                mk = which_map.get(sec, [])
                present = "YES" if found_map.get(sec, False) else "NO"
                sec_rows.append({
                    "PDF_FILE_NAME": pdf.name,
                    "SECTION": sec,
                    "MATCHED_KEYWORDS": ", ".join(mk),
                    "PRESENT_YN": present,
                    "REMARKS": "",
                })

            # Missing pages detection (once per file)
            page_issue_triggered = False
            page_issue_remarks = ""
            if check_pages:
                has_issue, rem = detect_missing_pages(prepared)
                page_issue_triggered = has_issue
                page_issue_remarks = rem

            # Rows for proposal_scan_result.csv
            for rule in issue_rules:
                rtype = str(rule["RuleType"])
                must_present = bool(rule.get("MustBePresent", True))
                scope = str(rule.get("Scope", "any")).lower()
                issue_id = rule.get("IssueID", "")
                issue_name = rule.get("IssueName", issue_id)
                matched_labels: List[str] = []
                pass_yn = "YES"
                remarks = ""

                if rtype == "missing_pages":
                    # PASS (YES) means no missing-page issue if MustBePresent=True
                    if must_present:
                        pass_yn = "NO" if page_issue_triggered else "YES"
                        remarks = page_issue_remarks or ""
                    else:
                        pass_yn = "YES" if page_issue_triggered else "NO"
                        remarks = page_issue_remarks or ""

                elif rtype == "missing_section":
                    sec_raw = str(rule.get("Section", "")).strip()
                    if not sec_raw:
                        typer.echo("[WARN] missing_section rule without 'Section' — skipped.")
                        continue

                    # normalize and map to canonical section key
                    sec_norm = _norm_key(sec_raw)
                    canon = section_key_map.get(sec_norm)
                    if not canon:
                        canon = sec_raw
                        typer.echo(f"[WARN] Section in Issues.xlsx not recognized: '{sec_raw}'. "
                                f"Expected one of {list(compiled_syn.keys())}.")

                    present = bool(found_map.get(canon, False))
                    pass_yn = "YES" if (present == must_present) else "NO"

                    # ⬇️ NEW: put matched headings into MATCHED_KEYWORDS
                    matched_labels = which_map.get(canon, []) if present else []

                    if not present:
                        remarks = "Section not detected"


                else:  # keyword
                    regex = str(rule.get("Regex", ""))
                    ok, labels = eval_keyword_rule(text, regex, must_present, scope, per_section_text)
                    pass_yn = "YES" if ok else "NO"
                    matched_labels = labels

                issue_rows.append({
                    "PDF_FILE_NAME": pdf.name,
                    "ISSUE_ID": issue_id,
                    "ISSUE_NAME": issue_name,
                    "RULE_TYPE": rtype,
                    "RULE_SCOPE": scope,
                    "MUST_BE_PRESENT": "YES" if must_present else "NO",
                    "MATCHED_KEYWORDS": ", ".join(matched_labels),
                    "PASS_YN": pass_yn,
                    "REMARKS": remarks,
                })

        except Exception as e:
            # Emit error rows for easier debugging (sections + one issue row)
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