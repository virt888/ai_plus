AI Plus - Proposal

————————————————————————
Python Command to run pdf scan 
Proposal scan
————————————————————————
Stpes for new env
python3 -m venv .venv

source .venv/bin/activate

# check version
python --version
# or
python3 --version


pip install -r requirements.txt

#NEW 
python scan_proposals.py \
  --input-dir ./in \
  --synonyms-csv rules/synonyms.csv \
  --issues-xlsx rules/issues_rules.xlsx \
  --out-sections-csv out/sections_presence.csv \
  --out-issues-csv out/proposal_scan_result.csv \
  --out-summary-csv out/summary_report.csv \
  --pages-limit 0 \
  --check-pages

