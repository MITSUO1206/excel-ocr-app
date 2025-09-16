# Excel抽出ツール（Streamlit）

Excelをドラッグ&ドロップして明細を抽出・集計するツールです。

## クイックスタート（Windows想定）

```bash
# 1) 取得
git clone https://github.com/<yourname>/<repo>.git
cd <repo>

# 2) 仮想環境 & 依存インストール
python -m venv .venv && .\.venv\Scripts\activate
pip install -r requirements.txt

# 3) 起動
streamlit run app_dragdrop_excel_reports.py
