# deal-report-automation

社内システムから当月の商談・案件データをCSV取得し、
個人別集計Excelを自動生成するスクリプトです。
OneDriveへのコピーとOutlookによるメール送信まで完全自動化します。
タスクスケジューラでの定期実行を想定。

## 機能

- Seleniumによるブラウザ自動操作（ログイン〜CSVダウンロードまで）
- pandas + openpyxlによる複数シート構成のExcel自動生成
- 商談・案件の個人別集計（新規・休眠・既存・受注率）
- xlwingsによる計算値変換・書式整形（枠線・フィルター・列幅）
- OneDriveへの自動コピー
- Outlookによる自動メール送信（送信アカウント指定対応）

## 効果

| 項目 | 改善前 | 改善後 |
|------|--------|--------|
| 作業時間 | 約1時間（手動） | 約5分（自動） |
| 作業方式 | 毎回手動操作 | タスクスケジューラで無人実行 |

## 必要環境

- Python 3.x
- Google Chrome
- Microsoft Outlook
- 以下のパッケージ

```
pip install selenium webdriver-manager pandas openpyxl xlwings python-dotenv pywin32
```

## セットアップ

1. `.env.example` をコピーして `.env` を作成
2. `.env` に各種設定を入力

```
LOGIN_ID=your_login_id
LOGIN_PASSWORD=your_password
TARGET_URL=https://your-system-url.example.com/
NAV_LABEL=your_nav_label
NAV_SUBLABEL_DEAL=your_deal_menu_label
NAV_SUBLABEL_PROJECT=your_project_menu_label
DOWNLOAD_FOLDER=\\your_server\path\to\folder
WORK_EXCEL_PATH=\\your_server\path\to\work.xlsx
TABLE_EXCEL_PATH=\\your_server\path\to\table.xlsx
OUTPUT_EXCEL_PATH=\\your_server\path\to\output.xlsx
ONEDRIVE_PATH=C:\Users\yourname\OneDrive\output.xlsx
MAIL_FROM=sender@example.com
MAIL_TO=recipient@example.com
MAIL_CC=cc@example.com
MAIL_SUBJECT=your_subject
ONEDRIVE_URL=https://your-onedrive-share-url
```

3. スクリプトを実行

```
python deal_report.py
```
