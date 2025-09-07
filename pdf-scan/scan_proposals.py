#!/usr/bin/env python3
import os, re, io, subprocess
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import fitz  # PyMuPDF
from PIL import Image
import pytesseract
import pandas as pd
import typer

app = typer.Typer(add_completion=False)

# -----------------------------
# Config (can be overridden)
# -----------------------------
# OCR
TESS_LANG = "eng"         # add "chi_tra+eng" if Traditional Chinese needed (requires traineddata installed)
TESS_CFG  = r"--oem 3 --psm 6"
DPI       = 150

# Default section synonyms (used for 'sections_presence.csv' and for "missing_section" issues)
DEFAULT_SECTION_SYNONYMS = {
    "Title":         [r"^\s*(proposed\s+topic|title of research|project title|title|topic)\b"],
    "Background":    [r"^\s*(background|overview|introduction)\b"],
    "Methodology":   [r"^\s*(methodology|research methods|materials and methods)\b"],
    "OutcomesValue": [r"^\s*(outcomes?\b|value\b|findings\b|analysis)\b"],
    "References":    [r"^\s*(references|bibliography|works cited)\b"],
}

# Built-in issue rules if your Issues.xlsx is missing or partial
# - missing_pages: always checked (detect gaps / odd-even only)
# - missing_section: for all default sections above
# - keyword rules: "timeline/schedule/gantt" and "budget/cost/expenditure" must be present somewhere
DEFAULT_ISSUE_RULES = [
    # Missing pages heuristics
    {"IssueID": "MISSING_PAGES", "IssueName": "Missing or odd/even-only pages",
     "RuleType": "missing_pages", "Regex": "", "MustBePresent": True, "Scope": "any"},

    # Missing section presence (Title, Background, Methodology, OutcomesValue, References)
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

    # Keyword presence rules (common FHDC flags)
    {"IssueID": "MUST_HAVE_TIMELINE", "IssueName": "Timeline/schedule not found",
     "RuleType": "keyword",
     "Regex": r"\b(timeline|schedule|gantt|milestone|workplan|time\s*line)\b",
     "MustBePresent": True, "Scope": "any"},
    {"IssueID": "MUST_HAVE_BUDGET", "IssueName": "Budget/cost not found",
     "RuleType": "keyword",
     "Regex": r"\b(budget|costs?|expenditure|expenses?|funding)\b",
     "MustBePresent": True, "Scope": "any"},
]

# -----------------------------
# Utilities
# -----------------------------
def compile_section_synonyms(mapping: Dict[str, List[str]]) -> Dict[str, List[re.Pattern]]:
    return {k: [re.compile(p, re.I | re.M) for p in v] for k, v in mapping.items()}

def which(cmd: str) -> Optional[str]:
    from shutil import which as _which
    return _which(cmd)

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
    if which("ocrmypdf"):
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
    return pdf  # fallback, will OCR pages on the fly

def page_to_text_via_ocr(page) -> str:
    mat = fitz.Matrix(DPI/72, DPI/72)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    return pytesseract.image_to_string(img, lang=TESS_LANG, config=TESS_CFG)

def pdf_to_text(pdf: Path, pages_limit: Optional[int]) -> str:
    """Extract text; if a page has no text, OCR that page."""
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
    """Return (has_issue, remarks). Detect gaps or odd/even only by printed numbers."""
    re_num = re.compile(r"\bpage\s+(\d+)\b(?:\s*of\s*(\d+))?", re.I)
    re_lone = re.compile(r"^\s*(\d{1,3})\s*$")
    nums: List[Optional[int]] = []
    with fitz.open(pdf) as doc:
        for i in range(doc.page_count):
            text = doc[i].get_text("text")
            candidates = []
            for line in text.splitlines():
                m = re_num.search(line)
                if m:
                    candidates.append(int(m.group(1)))
                    continue
                m2 = re_lone.match(line)
                if m2:
                    candidates.append(int(m2.group(1)))
            nums.append(min(candidates, key=lambda x: abs(x - (i + 1))) if candidates else None)
    present = [n for n in nums if n is not None]
    if not present:
        return (False, "")  # cannot conclude
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
                lbl = m.group(1) if m.lastindex else m.group(0)
                labels.add(lbl.strip())
                hit = True
        found_map[section] = hit
        which_map[section] = sorted(labels)
    return found_map, which_map

# -----------------------------
# Issues loader
# -----------------------------
def load_issues_xlsx(path: Optional[Path]) -> pd.DataFrame:
    if path and path.exists():
        try:
            df = pd.read_excel(path)
            # normalize columns
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception:
            pass
    # empty DataFrame if not provided
    return pd.DataFrame()

def build_issue_rules(df: pd.DataFrame, default_sections: List[str]) -> List[dict]:
    rules: List[dict] = []
    # Always include missing_pages + missing_section defaults (unless user overrides with explicit RuleType rows)
    has_missing_pages = False
    has_missing_section = False

    if not df.empty:
        for _, row in df.iterrows():
            rule = {}
            # Basic fields with defaults
            rule["IssueID"] = str(row.get("IssueID") or f"ISSUE_{_+1}").strip()
            rule["IssueName"] = str(row.get("IssueName") or "").strip() or rule["IssueID"]
            rt = str(row.get("RuleType") or "").strip().lower() or "keyword"
            rule["RuleType"] = rt  # keyword / missing_pages / missing_section
            rule["Regex"] = str(row.get("Regex") or row.get("Keyword") or "").strip()
            rule["MustBePresent"] = bool(row.get("MustBePresent")) if pd.notna(row.get("MustBePresent")) else True
            rule["Scope"] = str(row.get("Scope") or "any").strip().lower()
            if rt == "missing_pages":
                has_missing_pages = True
            if rt == "missing_section":
                has_missing_section = True
                sec = str(row.get("Section") or "").strip()
                if not sec:
                    continue  # invalid missing_section rule
                rule["Section"] = sec
            rules.append(rule)

    # Add built-in defaults if not overridden
    if not has_missing_pages:
        rules.append([r for r in DEFAULT_ISSUE_RULES if r["RuleType"] == "missing_pages"][0])

    # Add a missing_section rule for each default section if the sheet didn’t define any
    if not has_missing_section:
        for sec in default_sections:
            rules.append({"IssueID": f"MISS_SEC_{sec}", "IssueName": f"Missing {sec} section",
                          "RuleType": "missing_section", "Section": sec,
                          "Regex": "", "MustBePresent": True, "Scope": "any"})

    # Add built-in keyword rules if the sheet didn’t provide any keyword rows
    if df.empty or not any((str(t).lower() == "keyword") for t in df.get("RuleType", [])):
        for r in DEFAULT_ISSUE_RULES:
            if r["RuleType"] == "keyword":
                rules.append(r)

    return rules

def eval_keyword_rule(text: str, regex: str, must_be_present: bool, scope: str,
                      section_texts: Optional[Dict[str, str]] = None) -> Tuple[bool, List[str]]:
    """Return (pass, matched_keywords). For 'scope=section:X', search only that section’s text."""
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
        # Must NOT be present (rare, but supported)
        return (len(matched) == 0, matched)

# -----------------------------
# CLI
# -----------------------------
@app.command()
def run(
    input_dir: Path = typer.Option(..., "--input-dir", "-i", help="Folder containing PDFs (recursively scanned)"),
    out_sections_csv: Path = typer.Option("sections_presence.csv", "--out-sections-csv", help="Output for section presence"),
    out_issues_csv: Path = typer.Option("issues_scan.csv", "--out-issues-csv", help="Output for issue rules"),
    synonyms_csv: Optional[Path] = typer.Option(None, "--synonyms-csv", "-s", help="Optional synonyms CSV (Section,KeywordRegex)"),
    issues_xlsx: Optional[Path] = typer.Option(None, "--issues-xlsx", "-x", help="Issues Excel to drive rule checks"),
    ocr_cache_dir: Path = typer.Option("ocr_out", "--ocr-cache-dir", help="Where to write OCR'ed PDFs if ocrmypdf is available"),
    pages_limit: int = typer.Option(0, "--pages-limit", help="0 = all pages; otherwise limit pages per PDF for speed"),
    check_pages: bool = typer.Option(True, "--check-pages/--no-check-pages", help="Heuristically flag missing/odd-even pages"),
):
    # 1) Load synonyms (sections)
    if synonyms_csv and synonyms_csv.exists():
        sdf = pd.read_csv(synonyms_csv)
        syn_map: Dict[str, List[str]] = {}
        for _, row in sdf.iterrows():
            syn_map.setdefault(str(row["Section"]).strip(), []).append(str(row["KeywordRegex"]).strip())
        compiled_syn = compile_section_synonyms(syn_map)
    else:
        compiled_syn = compile_section_synonyms(DEFAULT_SECTION_SYNONYMS)

    # 2) Build issue rules (merge sheet + defaults)
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

    # Buffers for outputs
    sec_rows = []
    issue_rows = []

    # 4) Process PDFs
    for idx, pdf in enumerate(pdfs, 1):
        typer.echo(f"[{idx}/{len(pdfs)}] Scanning {pdf.name} ...")
        try:
            prepared = ensure_ocr_pdf(pdf, Path(ocr_cache_dir))
            text = pdf_to_text(prepared, pages_limit if pages_limit > 0 else None)

            # Section presence
            found_map, which_map = find_section_presence(text, compiled_syn)

            # Optionally build per-section text blocks to support scoped keyword rules
            per_section_text: Dict[str, str] = {}
            for sec_name, patterns in compiled_syn.items():
                # naive approach: if any heading appears, take +/- a window around it (simple heuristic)
                # For now, we just search whole text; you can refine to split by headings if needed.
                per_section_text[sec_name] = text

            # Sections CSV rows
            for sec in compiled_syn.keys():
                mk = which_map.get(sec, [])
                present = "YES" if found_map.get(sec, False) else "NO"
                remarks = ""
                sec_rows.append({
                    "PDF_FILE_NAME": pdf.name,
                    "SECTION": sec,
                    "MATCHED_KEYWORDS": ", ".join(mk),
                    "PRESENT_YN": present,
                    "REMARKS": remarks
                })

            # Issues: missing_pages (if enabled)
            page_issue_triggered = False
            page_issue_remarks = ""
            if check_pages:
                has_issue, rem = detect_missing_pages(prepared)
                page_issue_triggered = has_issue
                page_issue_remarks = rem

            # Issues CSV rows
            for rule in issue_rules:
                rtype = rule["RuleType"]
                must_present = bool(rule.get("MustBePresent", True))
                scope = str(rule.get("Scope", "any")).lower()
                issue_id = rule.get("IssueID", "")
                issue_name = rule.get("IssueName", issue_id)
                matched_labels: List[str] = []
                pass_yn = "YES"
                remarks = ""

                if rtype == "missing_pages":
                    # Interpret: PASS if no missing-page issue when presence is expected (MustBePresent=True)
                    # We flip semantics so PASS=good, FAIL=issue remains.
                    if must_present:
                        # We "expect" pages to be okay: if issue is triggered → FAIL
                        pass_yn = "NO" if page_issue_triggered else "YES"
                        remarks = page_issue_remarks or ""
                    else:
                        # Must not be present (rare)
                        pass_yn = "YES" if page_issue_triggered else "NO"
                        remarks = page_issue_remarks or ""

                elif rtype == "missing_section":
                    sec = str(rule.get("Section", "")).strip()
                    if not sec:
                        continue
                    present = bool(found_map.get(sec, False))
                    # If section must be present, PASS when found; else PASS when absent
                    pass_yn = "YES" if (present == must_present) else "NO"
                    if not present:
                        remarks = "Section not detected"

                else:  # keyword rules
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
                    "REMARKS": remarks
                })

        except Exception as e:
            # On error, still emit rows indicating failure
            for sec in compiled_syn.keys():
                sec_rows.append({
                    "PDF_FILE_NAME": pdf.name,
                    "SECTION": sec,
                    "MATCHED_KEYWORDS": "",
                    "PRESENT_YN": "ERROR",
                    "REMARKS": f"{type(e).__name__}: {e}"
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
                "REMARKS": f"{type(e).__name__}: {e}"
            })

    # 5) Write outputs
    pd.DataFrame(sec_rows).to_csv(out_sections_csv, index=False)
    pd.DataFrame(issue_rows).to_csv(out_issues_csv, index=False)
    typer.secho(f"Wrote {out_sections_csv} and {out_issues_csv} (files scanned: {len(pdfs)})", fg=typer.colors.GREEN)

if __name__ == "__main__":
    app()