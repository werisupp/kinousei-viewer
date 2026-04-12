#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
機能性表示食品 ビューワー生成スクリプト
対応フォーマット: 統合データ(114列) + SR情報抽出 シート（2025年4月以降）

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

xlsx_path = sorted(xlsx_files, key=os.path.getmtime, reverse=True)[0]
print(f"読み込み中: {os.path.basename(xlsx_path)}")

# ---- Excel読み込み ----
try:
    print("Excelを読み込み中（少し時間がかかります…）")
    xls = pd.ExcelFile(xlsx_path)
    sheet_names = xls.sheet_names
    print(f"  シート一覧: {sheet_names}")

    # 統合データシートを探す（旧フォーマット: 1枚目シートにも対応）
    MAIN_SHEET = None
    for cand in ["統合データ", "届出情報"]:
        if cand in sheet_names:
            MAIN_SHEET = cand
            break
    if MAIN_SHEET is None:
        MAIN_SHEET = sheet_names[0]

    df = pd.read_excel(xls, sheet_name=MAIN_SHEET, dtype=str, header=0)
    print(f"  [{MAIN_SHEET}] {len(df):,} 件、{len(df.columns)} 列")

    # SR情報抽出シートを探す
    SR_SHEET = "SR情報抽出" if "SR情報抽出" in sheet_names else None
    df_sr = None
    if SR_SHEET:
        df_sr = pd.read_excel(xls, sheet_name=SR_SHEET, dtype=str, header=0)
        df_sr = df_sr.fillna("")
        print(f"  [{SR_SHEET}] {len(df_sr):,} 件、{len(df_sr.columns)} 列")

except Exception as e:
    print(f"ERROR: Excelの読み込みに失敗しました: {e}")
    input("Enterで閉じる...")
    sys.exit(1)

df = df.fillna("")
all_cols = list(df.columns)

# ---- 統合データ: 全列をテーブル表示 ----
table_cols_exist = all_cols

# 表示ラベル（長い列名を短縮表示するためのマッピング）
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

# ---- SR情報をグルーピング（届出番号をキーに） ----
sr_map = {}
if df_sr is not None and "届出番号" in df_sr.columns:
    sr_cols = [c for c in df_sr.columns if c != "届出番号"]
    for nonum, grp in df_sr.groupby("届出番号"):
        rows = []
        for _, row in grp.iterrows():
            entry = {c: str(row[c]) for c in sr_cols if str(row.get(c, "")).strip() and str(row.get(c, "")) != "nan"}
            if entry:
                rows.append(entry)
        if rows:
            sr_map[str(nonum)] = rows

# ---- SR一覧テーブル用データ（フラット化）----
sr_table_records = []
sr_table_cols = []
if df_sr is not None:
    sr_table_cols = list(df_sr.columns)
    for _, row in df_sr.iterrows():
        rec = {c: str(row[c]) for c in sr_table_cols}
        search_text = " ".join(str(row[c]) for c in sr_table_cols).lower()
        sr_table_records.append({"t": rec, "s": search_text})

# ---- JSONデータ変換 ----
print("JSONデータに変換中…")
search_cols = [
    "商品名",
    "法人名",
    "機能性関与成分名",
    "機能性関与成分名.1",
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
    nonum = str(row.get("届出番号", "")).strip()
    sr_rows = sr_map.get(nonum, [])
    records.append({"t": table_data, "s": search_text, "d": detail_data, "sr": sr_rows})

json_data = json.dumps(records, ensure_ascii=False)
table_cols_json = json.dumps(table_cols_labels, ensure_ascii=False)
sr_records_json = json.dumps(sr_table_records, ensure_ascii=False)
sr_cols_json = json.dumps(sr_table_cols, ensure_ascii=False)
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
has_sr_json = "true" if sr_map else "false"

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
    .replace("{{SR_RECORDS_JSON}}", sr_records_json)
    .replace("{{SR_COLS_JSON}}", sr_cols_json)
    .replace("{{KUBUN_LABEL}}", kubun_label)
    .replace("{{STATUS_LABEL}}", status_label)
    .replace("{{HAS_SR}}", has_sr_json)
)

out_path = os.path.join(script_dir, "viewer.html")
with open(out_path, "w", encoding="utf-8") as f:
    f.write(html)

print(f"✅ viewer.html を生成しました！")
print(f"   保存先: {out_path}")
print("ブラウザを開いています…")
webbrowser.open("file:///" + out_path.replace(os.sep, "/"))
print("完了！")
