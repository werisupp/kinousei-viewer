#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
機能性表示食品 ビューワー生成スクリプト
同じフォルダの .xlsx を読み込み viewer.html を生成してブラウザで開く

使い方:
  Mac:     python3 build_viewer.py
  Windows: python  build_viewer.py
"""
import sys
import os
import glob
import json
import webbrowser
from datetime import datetime

# ---- ライブラリ確認 ----
try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas がありません。")
    print("以下を実行してください:")
    print("  pip3 install pandas openpyxl  (Mac)")
    print("  pip  install pandas openpyxl  (Windows)")
    input("Enterで閉じる...")
    sys.exit(1)

try:
    import openpyxl  # noqa: F401
except ImportError:
    print("ERROR: openpyxl がありません。")
    print("以下を実行してください:")
    print("  pip3 install openpyxl  (Mac)")
    print("  pip  install openpyxl  (Windows)")
    input("Enterで閉じる...")
    sys.exit(1)

# ---- Excelファイルを探す ----
script_dir = os.path.dirname(os.path.abspath(__file__))
xlsx_files = glob.glob(os.path.join(script_dir, "*.xlsx"))

if not xlsx_files:
    print("ERROR: .xlsx ファイルが見つかりません。")
    print(f"このフォルダに置いてください: {script_dir}")
    input("Enterで閉じる...")
    sys.exit(1)

# 最終更新日時が最新のファイルを選択
xlsx_path = sorted(xlsx_files, key=os.path.getmtime, reverse=True)[0]
print(f"読み込み中: {os.path.basename(xlsx_path)}")

# ---- Excel読み込み ----
try:
    print("Excelを読み込み中（少し時間がかかります…）")
    df = pd.read_excel(xlsx_path, sheet_name=0, dtype=str, header=0)
    print(f"  {len(df):,} 件、{len(df.columns)} 列")
except Exception as e:
    print(f"ERROR: Excelの読み込みに失敗しました: {e}")
    input("Enterで閉じる...")
    sys.exit(1)

df = df.fillna("")
all_cols = list(df.columns)

# ---- テーブル表示列の設定 ----
# 新フォーマット（2025年4月以降）対応
TABLE_COLS_PREFER = [
    "届出番号",
    "届出日",
    "法人名",
    "商品名",
    "名称",
    "機能性関与成分名",
    "表示しようとする機能性",
    "食品の区分",
    "（届出日から60日経過した場合）販売状況",
    "販売開始予定日",
]

COL_LABELS = {
    "届出番号": "届出番号",
    "届出日": "届出日",
    "撤回日": "撤回日",
    "変更日": "変更日",
    "法人名": "法人名",
    "商品名": "商品名",
    "名称": "名称",
    "機能性関与成分名": "機能性関与成分",
    "表示しようとする機能性": "機能性",
    "食品の区分": "食品区分",
    "（届出日から60日経過した場合）販売状況": "販売状況",
    "販売開始予定日": "販売開始予定日",
    "当該製品が想定する主な対象者（疾病に罹患している者、未成年者、妊産婦（妊娠を計画している者を含む。）及び授乳婦を除く。）": "想定対象者",
    "情報開示するウェブサイトのＵＲＬ": "WebサイトURL",
}

table_cols_exist = [c for c in TABLE_COLS_PREFER if c in all_cols]
if len(table_cols_exist) < 3:
    # 想定列が見つからない場合は先頭10列を使用
    table_cols_exist = all_cols[:10]
    print("注意: 想定されている列名が見つかりませんでした。先頭10列を使用します。")

table_cols_labels = [COL_LABELS.get(c, c) for c in table_cols_exist]

# ---- フィルター選択肢の収集 ----
kubun_col = next((c for c in all_cols if "食品の区分" in c), None)
status_col = next((c for c in all_cols if "販売状況" in c), None)

kubun_values = (
    sorted([v for v in df[kubun_col].unique().tolist() if str(v).strip()])
    if kubun_col else []
)
status_values = (
    sorted([v for v in df[status_col].unique().tolist() if str(v).strip()])
    if status_col else []
)

# ---- JSONデータ変換 ----
print("JSONデータに変換中…")
search_cols = [
    "商品名",
    "法人名",
    "機能性関与成分名",
    "表示しようとする機能性",
    "届出番号",
    "名称",
]
records = []
for _, row in df.iterrows():
    table_data = {COL_LABELS.get(c, c): str(row[c]) for c in table_cols_exist}
    search_text = " ".join(
        str(row[c]) for c in search_cols if c in all_cols
    ).lower()
    detail_data = {
        c: str(row[c])
        for c in all_cols
        if str(row[c]).strip() and str(row[c]) != "nan"
    }
    records.append({"t": table_data, "s": search_text, "d": detail_data})

json_data = json.dumps(records, ensure_ascii=False)
table_cols_json = json.dumps(table_cols_labels, ensure_ascii=False)
kubun_opts = "\n".join(
    f'<option value="{v}">{v}</option>' for v in kubun_values
)
status_opts = "\n".join(
    f'<option value="{v}">{v}</option>' for v in status_values
)
update_time = datetime.now().strftime("%Y年%m月%d日 %H:%M")
source_file = os.path.basename(xlsx_path)
kubun_label = COL_LABELS.get("食品の区分", "食品区分")
status_label = COL_LABELS.get(
    "（届出日から60日経過した場合）販売状況", "販売状況"
)

# ---- テンプレート読み込み ----
tpl_path = os.path.join(script_dir, "_template.html")
if not os.path.exists(tpl_path):
    print(f"ERROR: _template.html が見つかりません: {tpl_path}")
    input("Enterで閉じる...")
    sys.exit(1)

with open(tpl_path, "r", encoding="utf-8") as f:
    template = f.read()

# ---- HTML生成（テンプレートの置換） ----
print("viewer.html を生成中…")
html = (
    template
    .replace("{{SOURCE_FILE}}", source_file)
    .replace("{{UPDATE_TIME}}", update_time)
    .replace("{{KUBUN_OPTIONS}}", kubun_opts)
    .replace("{{STATUS_OPTIONS}}", status_opts)
    .replace("{{ALL_DATA_JSON}}", json_data)
    .replace("{{TABLE_COLS_JSON}}", table_cols_json)
    .replace("{{KUBUN_LABEL}}", kubun_label)
    .replace("{{STATUS_LABEL}}", status_label)
)

out_path = os.path.join(script_dir, "viewer.html")
with open(out_path, "w", encoding="utf-8") as f:
    f.write(html)

print(f"✅ viewer.html を生成しました！")
print(f"   保存先: {out_path}")
print("ブラウザを開いています…")
webbrowser.open("file:///" + out_path.replace(os.sep, "/"))
print("完了！")
