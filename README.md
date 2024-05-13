# Excel to Word Report Generator

## 描述
這個 Python 腳本 (`excel2word.py`) 自動從 Excel 檔案讀取數據並生成格式化的 Word 報告文檔。它特別適用於自動化測試報告的生成。

## 功能
- 讀取 Excel 測試數據
- 填充 Word 文檔表格
- 生成與 Excel 同名的 Word 文檔

## 快速開始
1. 直接下載並使用我們提供的可執行文件。請訪問[釋放頁面](https://github.com/tangchi1215/TestCase/releases/tag/v1.0.1)下載最新版本。
2. 並下載[excel模板](https://github.com/tangchi1215/TestCase/tree/master/templates)，開始撰寫測試案例。

## 安裝
1. Clone Repository 或 下載 Source Code：
   ```bash
   git clone https://example.com/your-repository.git](https://github.com/tangchi1215/TestCase.git)
   ```
2. 安裝所需的依賴：
   ```bash
   pip install -r requirements.txt
   ```

## 使用方法
1. 使用template.xlsx撰寫你的測試案例，可改成你想要的檔名
2. 運行腳本：
   ```bash
   python excel2word.py
   ```
3. 按照終端中的提示操作。

## 打包成 EXE
如果你想自己打包成 `.exe`，請遵循以下步驟：
1. 安裝 PyInstaller：
   ```bash
   pip install pyinstaller
   ```
2. 使用 PyInstaller 打包腳本：
   ```bash
   pyinstaller --onefile excel2word.py
   ```
## 更新requirements.txt
1. 使用pip freeze
   ```bash
   pip freeze > requirements.txt
   ```

   
