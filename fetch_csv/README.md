# fetch_csv — CSVダウンロード・マスタ更新スクリプト

CAA（消費者庁）の機能性表示食品届出情報データベースから
**「前日までの全届出の全項目出力(CSV出力)」** を自動でダウンロードし、
マスタExcelファイルを最新状態に保つスクリプトです。

## フォルダ構成

```
kinousei-viewer/
└── fetch_csv/
    ├── fetch_csv.py        ← メインスクリプト
    ├── README.md           ← このファイル
    ├── downloads/          ← ダウンロードしたCSVファイルの保管場所（自動生成）
    │   └── *.csv
    └── master.xlsx         ← 統合マスタファイル（自動生成）
```

## 必要ライブラリのインストール

```bash
# Mac / Linux
pip3 install requests playwright pandas openpyxl
playwright install chromium

# Windows
pip install requests playwright pandas openpyxl
playwright install chromium
```

## 実行方法

```bash
# リポジトリのルートから実行
python3 fetch_csv/fetch_csv.py

# または fetch_csv/ フォルダに移動して実行
cd fetch_csv
python3 fetch_csv.py
```

## 処理の流れ

```
[1/3] PlaywrightでCAAページにアクセス
         ↓
      「前日までの全届出の全項目出力」ボタンをクリック
         ↓
      CSVを fetch_csv/downloads/ に保存
         ↓
[2/3] downloads/ 内の全CSVを統合（重複除去）
         ↓
[3/3] master.xlsx を作成 or 差分更新
```

## マスタファイルの更新ロジック

| 状況 | 動作 |
|------|------|
| `master.xlsx` が存在しない（初回） | 統合データで新規作成 |
| 届出番号が新しい行 | マスタに追加 |
| 届出番号が一致するが他の列が異なる | その行を最新情報で上書き |
| 届出番号・全列が一致（変化なし） | 何もしない |

## ダウンロードファイルの重複処理

同名かつ同一内容のCSVが既に `downloads/` に存在する場合は
ダウンロードをスキップします。
内容が異なる場合はタイムスタンプ付きの別名で保存します。

## 注意事項

- `playwright install chromium` の実行が必要です
- ヘッドレスブラウザ（Chromium）を使用するため、初回は数分かかる場合があります
- CAAサイトの構造変更によりボタン検出に失敗することがあります。
  その場合は `fetch_csv.py` 内の `BUTTON_TEXT` 変数を確認してください
- CSVのエンコーディングはCP932（Shift-JIS）とUTF-8-sigの両方に対応しています
