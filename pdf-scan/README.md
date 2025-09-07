AI Plus - Proposal

————————————————————————
Python Command to run pdf scan 
Proposal scan
————————————————————————
Stpes for new env
/opt/homebrew/bin/python3.12 -m venv .venv

source .venv/bin/activate

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
  --pages-limit 0 \
  --check-pages

Or
python scan_proposals.py \
  --input-dir ./in \
  --synonyms-csv rules/synonyms.csv \
  --issues-xlsx rules/issues_rules.xlsx \
  --out-sections-csv out/sections_presence.csv \
  --out-issues-csv out/proposal_scan_result.csv \
  --with-defaults \
  --pages-limit 0 \
  --check-pages




#OLD
# (venv active, deps installed, tesseract/ocrmypdf installed)
python scan_proposals.py \
  --input-dir ./in \
  --out-sections-csv sections_presence.csv \
  --out-issues-csv proposal_scan_result.csv \
  --synonyms-csv synonyms.csv \
  --issues-xlsx Issues.xlsx \
  --pages-limit 0 \
  --check-pages

# Reset or quit envdeactivate
