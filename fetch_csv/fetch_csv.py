#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fetch_csv.py
機能性表示食品届出情報データベースから
「前日までの全届出の全項目出力(CSV)」をダウンロードし、
ダウンロードフォルダ（fetch_csv/downloads/）に保存する。

その後、全CSVを統合して:
  - 初回実行 → fetch_csv/master.xlsx を新規作成
  - 2回目以降 → 「届出番号」キーで差分を検出し、master.xlsx を上書き更新

使い方:
  pip install requests playwright openpyxl pandas
  playwright install chromium
  python fetch_csv/fetch_csv.py
"""
import os
import sys
import glob
import time
import hashlib
from datetime import datetime
from pathlib import Path

# ライブラリ確認 ----------------------------------------------------------------
def _require(pkg, install_hint):
    try:
        __import__(pkg)
    except ImportError:
        print(f"ERROR: {pkg} がありません。")
        print(f"  {install_hint}")
        sys.exit(1)

_require("requests",   "pip install requests")
_require("playwright", "pip install playwright && playwright install chromium")
_require("pandas",     "pip install pandas openpyxl")
_require("openpyxl",   "pip install openpyxl")

import requests  # noqa: E402
import pandas as pd  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

# パス設定 ---------------------------------------------------------------------
SCRIPT_DIR   = Path(__file__).resolve().parent          # fetch_csv/
DOWNLOAD_DIR = SCRIPT_DIR / "downloads"                 # fetch_csv/downloads/
MASTER_PATH  = SCRIPT_DIR / "master.xlsx"               # fetch_csv/master.xlsx
TARGET_URL   = "https://www.fld.caa.go.jp/caaks/s/cssc01/"
BUTTON_TEXT  = "前日までの全届出の全項目出力"
KEY_COL      = "届出番号"                                # マスタ更新のキー列

DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

# ===========================================================================
# Step 1: PlaywrightでCSVダウンロード
# ===========================================================================
def download_csv() -> Path | None:
    """
    CAA届出情報DBページにアクセスし、
    「前日までの全届出の全項目出力(CSV出力)」ボタンを押して
    CSVファイルをダウンロードする。
    保存したファイルのパスを返す。
    """
    print("[1/3] ページにアクセスしてCSVをダウンロードします...")
    print(f"  URL: {TARGET_URL}")

    saved_path: Path | None = None

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page    = context.new_page()

        page.goto(TARGET_URL, timeout=60_000)
        page.wait_for_load_state("networkidle", timeout=30_000)

        # ボタンを探す（テキストに「前日までの全届出の全項目出力」を含む要素）
        btn = page.get_by_text(BUTTON_TEXT, exact=False).first
        if not btn:
            print(f"ERROR: ボタン '{BUTTON_TEXT}' が見つかりません。")
            browser.close()
            return None

        print(f"  ボタン発見: '{btn.inner_text().strip()}' → クリック")

        with page.expect_download(timeout=120_000) as dl_info:
            btn.click()

        download = dl_info.value
        suggested = download.suggested_filename or f"kinousei_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

        dest = DOWNLOAD_DIR / suggested
        # 同名ファイルが既にある場合はハッシュでスキップ判定するため、一旦一時保存
        tmp  = DOWNLOAD_DIR / ("_tmp_" + suggested)
        download.save_as(str(tmp))

        # 同内容のファイルがすでに存在するか確認
        if dest.exists():
            if _file_hash(tmp) == _file_hash(dest):
                print(f"  既に同じ内容のファイルが存在します: {dest.name} (スキップ)")
                tmp.unlink()
                saved_path = dest
            else:
                # 異なる内容 → タイムスタンプ付きで保存
                ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
                stem = dest.stem
                suf  = dest.suffix
                dest = DOWNLOAD_DIR / f"{stem}_{ts}{suf}"
                tmp.rename(dest)
                print(f"  保存しました: {dest.name}")
                saved_path = dest
        else:
            tmp.rename(dest)
            print(f"  保存しました: {dest.name}")
            saved_path = dest

        browser.close()

    return saved_path


def _file_hash(path: Path) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


# ===========================================================================
# Step 2: 全CSVを統合
# ===========================================================================
def load_all_csvs() -> pd.DataFrame:
    """
    downloads/ フォルダ内の全CSVを読み込み、統合した DataFrame を返す。
    1行目を見出し行として扱い、2行目以降を結合する。
    重複行（全列一致）は除去する。
    """
    csv_files = sorted(DOWNLOAD_DIR.glob("*.csv"))
    if not csv_files:
        print("ERROR: downloads/ フォルダにCSVファイルが見つかりません。")
        sys.exit(1)

    print(f"[2/3] {len(csv_files)} 個のCSVを統合します...")
    frames = []
    for f in csv_files:
        try:
            df = pd.read_csv(
                f,
                encoding="cp932",   # CAA公開CSVはShift-JIS(CP932)が多い
                dtype=str,
                header=0,
                on_bad_lines="skip",
            )
            frames.append(df)
            print(f"  読み込み: {f.name}  ({len(df):,} 件)")
        except UnicodeDecodeError:
            # UTF-8 フォールバック
            df = pd.read_csv(
                f,
                encoding="utf-8-sig",
                dtype=str,
                header=0,
                on_bad_lines="skip",
            )
            frames.append(df)
            print(f"  読み込み(UTF-8): {f.name}  ({len(df):,} 件)")
        except Exception as e:
            print(f"  警告: {f.name} の読み込みをスキップしました ({e})")

    if not frames:
        print("ERROR: 読み込めるCSVがありませんでした。")
        sys.exit(1)

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.fillna("")
    # 完全重複を除去
    combined = combined.drop_duplicates()
    print(f"  統合後: {len(combined):,} 件 / {len(combined.columns)} 列")
    return combined


# ===========================================================================
# Step 3: マスタファイルの作成 or 更新
# ===========================================================================
def update_master(new_df: pd.DataFrame):
    """
    初回 → master.xlsx を新規作成
    2回目以降:
      - 届出番号が新規 → 追加
      - 届出番号が一致するが他列が異なる → 上書き
    """
    if not MASTER_PATH.exists():
        # 初回
        print(f"[3/3] master.xlsx を新規作成します: {MASTER_PATH}")
        new_df.to_excel(str(MASTER_PATH), index=False, engine="openpyxl")
        print(f"  ✅ 作成完了: {len(new_df):,} 件")
        return

    # 2回目以降
    print("[3/3] master.xlsx と差分を確認して更新します...")
    try:
        master_df = pd.read_excel(str(MASTER_PATH), dtype=str, header=0)
        master_df = master_df.fillna("")
    except Exception as e:
        print(f"ERROR: master.xlsx の読み込みに失敗しました: {e}")
        sys.exit(1)

    # 見出し列をチェック
    if KEY_COL not in master_df.columns:
        print(f"WARNING: master.xlsx に '{KEY_COL}' 列が見つかりません。")
        print("  全データを上書きします。")
        new_df.to_excel(str(MASTER_PATH), index=False, engine="openpyxl")
        print(f"  ✅ 上書き完了: {len(new_df):,} 件")
        return

    if KEY_COL not in new_df.columns:
        print(f"WARNING: 新しいCSVに '{KEY_COL}' 列が見つかりません。")
        print("  全データを上書きします。")
        new_df.to_excel(str(MASTER_PATH), index=False, engine="openpyxl")
        print(f"  ✅ 上書き完了: {len(new_df):,} 件")
        return

    # 共通列を揃える（列の追加にも対応）
    all_cols = list(dict.fromkeys(list(master_df.columns) + list(new_df.columns)))
    for col in all_cols:
        if col not in master_df.columns:
            master_df[col] = ""
        if col not in new_df.columns:
            new_df[col] = ""
    master_df = master_df[all_cols]
    new_df    = new_df[all_cols]

    # インデックスを届出番号に
    master_indexed = master_df.set_index(KEY_COL)
    new_indexed    = new_df.set_index(KEY_COL)

    added   = 0
    updated = 0

    for key, new_row in new_indexed.iterrows():
        if key not in master_indexed.index:
            # 新規行
            master_indexed = pd.concat([
                master_indexed,
                new_row.to_frame().T
            ])
            added += 1
        else:
            # 既存行 → 差分チェック
            existing = master_indexed.loc[key]
            # 同じキーが複数行ある場合は最初の1件を対象にする
            if isinstance(existing, pd.DataFrame):
                existing = existing.iloc[0]
            cols_to_check = [c for c in all_cols if c != KEY_COL]
            if not existing[cols_to_check].equals(new_row[cols_to_check]):
                master_indexed.loc[key] = new_row
                updated += 1

    result = master_indexed.reset_index()
    result.to_excel(str(MASTER_PATH), index=False, engine="openpyxl")

    print(f"  新規追加: {added:,} 件")
    print(f"  上書き更新: {updated:,} 件")
    print(f"  ✅ master.xlsx 更新完了: {len(result):,} 件")


# ===========================================================================
# メイン
# ===========================================================================
def main():
    print("=" * 60)
    print("機能性表示食品届出情報 CSV取得・マスタ更新スクリプト")
    print(f"  実行日時: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}")
    print("=" * 60)

    # Step 1: CSVダウンロード
    saved = download_csv()
    if saved is None:
        print("CSV のダウンロードに失敗しました。処理を中断します。")
        sys.exit(1)

    # Step 2: 全CSV統合
    combined = load_all_csvs()

    # Step 3: マスタ更新
    update_master(combined)

    print()
    print("🎉 すべての処理が完了しました。")
    print(f"   マスタファイル: {MASTER_PATH}")


if __name__ == "__main__":
    main()
