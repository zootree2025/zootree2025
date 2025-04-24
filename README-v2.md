# PDF 轉 DOCX 轉換器 (V2)

`pdf2docV2.py` 是一個基於 Python 的圖形應用程式，允許使用者將 PDF 文件轉換為 DOCX 格式。該工具採用 `Tkinter` 創建直觀的圖形使用者介面 (GUI)，並使用 `pdf2docx` 庫進行文件轉換。

## 功能特性

- **友好的使用者介面**：透過 GUI 輕鬆選擇輸入 PDF 文件並指定輸出 DOCX 文件的路徑。
- **進度指示**：包含進度條和動畫視覺反饋，顯示轉換過程進度。
- **錯誤處理**：在轉換失敗時顯示錯誤訊息。
- **多執行緒處理**：在執行長時間任務時，確保應用程式保持回應。

## 使用前準備

在使用此腳本之前，請確保已安裝以下項目：

- Python 3.x
- 所需的 Python 庫：
  - `tkinter` (Python 預設已安裝)
  - `pdf2docx`

您可以使用以下指令安裝 `pdf2docx`：
```bash
pip install pdf2docx
