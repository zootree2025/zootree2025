PDF 轉 DOCX 轉換器
這是一個使用 Python 開發的 GUI 應用程式，可以將 PDF 文件轉換為 DOCX 格式。該工具基於 tkinter 和 pdf2docx 庫構建，提供了簡單易用的圖形化界面。

功能特性
輕鬆選擇 PDF 文件並指定 DOCX 輸出路徑
實時顯示轉換進度
錯誤提示與成功通知
安裝
在開始使用之前，請確保已安裝以下依賴項：

Python 3.x
tkinter（通常隨 Python 自帶）
pdf2docx 庫
您可以使用以下指令安裝 pdf2docx 庫：

bash
pip install pdf2docx
使用方法
下載並運行 pdf2doc.py：

bash
python pdf2doc.py
在應用程式中：

點擊「瀏覽」按鈕選擇 PDF 文件。
點擊「瀏覽」按鈕選擇 DOCX 輸出路徑。
點擊「開始轉換」按鈕開始文件轉換。
完成後，您將在指定的輸出路徑中找到轉換完成的 DOCX 文件。

目錄結構
plaintext
.
├── pdf2doc.py  # 主程式文件
注意事項
請確保 PDF 文件格式正確，否則可能會導致轉換失敗。
如果遇到任何錯誤訊息，請檢查是否安裝了所有必要的依賴項。
技術細節
界面框架: 使用 tkinter 構建。
轉換核心: 使用 pdf2docx 庫進行 PDF 到 DOCX 的轉換。
貢獻
如果您有任何改進建議，歡迎提交 Pull Request 或創建 Issue。

授權
本項目採用 MIT 授權許可，詳情請參閱 LICENSE。
