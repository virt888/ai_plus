Windows 安裝與執行指南

此說明文件提供在 Windows 系統上安裝 Python、設定虛擬環境並執行提案掃描工具 (scan_proposals.py) 的步驟。請按順序完成以下操作。

1. 安裝 Python
	1.	訪問 Python 官網下載頁：https://www.python.org/downloads/windows/。
	2.	點選最新版本的 Python（建議使用 Python 3.9 以上）。下載適用於 Windows 的執行檔（例如 python-3.10.***.exe）。
	3.	執行下載好的安裝程式，務必勾選「Add Python to PATH」，以便在命令提示符中直接使用 python 指令。
	4.	若沒有系統管理員權限，可以在安裝向導中選擇「Install for current user only」（僅安裝至當前使用者），避免需要管理員權限。接著點擊「Install Now」完成安裝。
	5.	安裝結束後，打開命令提示符（cmd 或 PowerShell），輸入：
	
	python --version

若顯示版本號，則表示 Python 安裝成功。

2. 建立與啟用虛擬環境

使用虛擬環境可以隔離套件，避免對系統其他專案造成影響。假設你已在專案根目錄下（包含 requirements.txt 和 scan_proposals.py 等檔案）。
	1.	打開命令提示符，執行：

	python -m venv .venv

這會在當前資料夾建立一個名為 .venv 的虛擬環境。

	2.	啟用虛擬環境：

	.\.venv\Scripts\activate


啟用成功後，提示符前方會出現 (.venv)，表示目前位於虛擬環境中。

3. 安裝相依套件

在虛擬環境啟用的情況下，使用 pip 安裝專案所需的套件：

pip install -r requirements.txt

此命令會根據 requirements.txt 安裝所有依賴程式，不需要管理員權限，因為套件會安裝在虛擬環境中。

4. 執行掃描腳本

安裝完依賴後，即可執行掃描腳本。以下命令示範如何掃描 ./in 目錄下的 PDF，並輸出結果：

python scan_proposals.py ^
  --input-dir .\in ^
  --synonyms-csv rules\synonyms.csv ^
  --issues-xlsx rules\issues_rules.xlsx ^
  --out-sections-csv out\sections_presence.csv ^
  --out-issues-csv out\proposal_scan_result.csv ^
  --out-summary-csv out\summary_report.csv ^
  --pages-limit 0 ^
  --check-pages

	•	^ 用於在 Windows 的命令提示符中換行（可將整條指令寫在一行，省略符號）。
	•	所有路徑使用反斜線 \；如果使用 PowerShell 或其他支援的終端機，也可以使用正斜線 /。

5. 可選：安裝 OCR 支援（Tesseract）

腳本中部分功能使用 pytesseract 進行文字辨識（OCR）。如果你需要處理掃描版 PDF，請安裝 Tesseract OCR：
	1.	前往 Tesseract OCR 下載頁 下載適用於 Windows 的安裝程式（通常為 .exe）。
	2.	依照指示完成安裝，安裝路徑通常預設在 C:\Program Files\Tesseract-OCR。
	3.	將安裝目錄加入系統環境變數 PATH，例如將 C:\Program Files\Tesseract-OCR 加入 PATH，或在程式碼中設定 pytesseract.pytesseract.tesseract_cmd。
	4.	安裝後可在命令提示符中輸入 tesseract --version 檢查是否成功。

6. 權限與常見問題
	•	安裝 Python 不一定需要管理員權限：在安裝向導中選擇「Install for current user only」即可將 Python 安裝在使用者目錄。
	•	安裝套件不需要管理員權限：透過虛擬環境安裝 pip 套件無需系統管理員權限。
	•	Visual C++ 依賴：某些套件（如 PyMuPDF）在 Windows 上可能需要 Microsoft Visual C++ Redistributable。若 pip install 過程中出現相關錯誤，請先安裝 Microsoft Visual C++ Redistributable。
	•	長路徑問題：在舊版 Windows，路徑長度可能受限制。若安裝路徑過長導致錯誤，可嘗試使用較短的目錄名。

完成以上步驟後，你就可以在 Windows 系統上順利執行提案掃描工具了。如遇到安裝或執行問題，請檢查安裝步驟是否正確或查閱相關套件的官方說明。
