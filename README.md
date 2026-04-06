# 機能性表示食品 届出データビューワー

消費者庁「機能性表示食品届出データベース」のExcelファイルを  
**ワンクリック**で見やすいHTMLビューワーに変換するツールです。

## 機能

- 🔍 リアルタイム全文検索（商品名・法人名・成分名・届出番号）
- 📂 食品区分・販売状況でフィルタリング
- ↕️ 各列ヘッダークリックでソート
- 📋 行クリックで全列データをポップアップ表示
- 🌙 ダークモード対応
- 📄 1ページ25/50/100/200件表示切替

## フォルダ構成

```
kinousei-viewer/
├── 更新してブラウザで開く.command   ← Mac用：ダブルクリックするファイル
├── 更新してブラウザで開く.bat       ← Windows用：ダブルクリックするファイル
├── build_viewer.py                  ← Python変換スクリプト
├── _template.html                   ← HTMLテンプレート
├── viewer.html                      ← 自動生成（最初は存在しない）
└── *.xlsx                           ← ここにExcelを置く
```

## セットアップ（初回のみ）

### 1. リポジトリをクローン

```bash
git clone https://github.com/werisupp/kinousei-viewer.git
cd kinousei-viewer
```

### 2. Pythonライブラリをインストール

```bash
pip install pandas openpyxl
```

### 3. Macの場合：commandファイルに実行権限を付与（初回1回だけ）

ターミナルで以下を実行：

```bash
cd ~/path/to/kinousei-viewer
chmod +x 更新してブラウザで開く.command
```

## 毎回の更新手順

1. [消費者庁](https://www.fld.caa.go.jp/caaks/s/cssc01/)からExcelをダウンロード
2. このフォルダ（`kinousei-viewer/`）にコピー
3. **Mac:** `更新してブラウザで開く.command` をダブルクリック  
   **Windows:** `更新してブラウザで開く.bat` をダブルクリック
4. ブラウザが自動で開き `viewer.html` が表示されます！

## よくある質問

**Q: ダブルクリックしても開かない（Mac）**  
A: ターミナルで `chmod +x 更新してブラウザで開く.command` を実行してください

**Q: 「開発元が未確認」と表示される（Mac）**  
A: Finder で右クリック →「開く」→「開く」をクリック

**Q: Excelが複数あると？**  
A: 更新日時が最新のファイルが自動で選ばれます

**Q: viewer.html を他の人に共有したい**  
A: `viewer.html` 単体をメール添付やUSBで渡せばそのまま動きます（インターネット不要）
