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
  python3 -m playwright install chromium
  python3 fetch_csv/fetch_csv.py
"""
import os
import sys
import glob
import hashlib
import time
import re
from datetime import datetime
from pathlib import Path
from typing import Optional, List  # Python 3.9 対応

# ライブラリ確認 ----------------------------------------------------------------
def _require(pkg, install_hint):
    try:
        __import__(pkg)
    except ImportError:
        print("ERROR: {} がありません。".format(pkg))
        print("  {}".format(install_hint))
        sys.exit(1)

_require("playwright", "pip install playwright && python3 -m playwright install chromium")
_require("pandas",     "pip install pandas openpyxl")
_require("openpyxl",   "pip install openpyxl")

import pandas as pd  # noqa: E402
from playwright.sync_api import sync_playwright, Download  # noqa: E402

# パス設定 ---------------------------------------------------------------------
SCRIPT_DIR   = Path(__file__).resolve().parent          # fetch_csv/
DOWNLOAD_DIR = SCRIPT_DIR / "downloads"                 # fetch_csv/downloads/
MASTER_PATH  = SCRIPT_DIR / "master.xlsx"               # fetch_csv/master.xlsx
TARGET_URL   = "https://www.fld.caa.go.jp/caaks/s/cssc01/"
BUTTON_TEXT  = "前日までの全届出の全項目出力"
KEY_COL      = "届出番号"                                # マスタ更新のキー列

# ダウンロード完了を待つ最大秒数（ボタンクリック後、最後のダウンロードから待機）
DOWNLOAD_TIMEOUT = 300  # 秒

# 最後のダウンロード発生から追加ダウンロードを待つ秒数
NEXT_DOWNLOAD_WAIT = 60  # 秒

# 保存するCSVのファイル名プレフィックス
FILE_PREFIX = "機能性表示食品全届出一覧"

DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

# UUID形式かどうか判定する正規表現
_UUID_RE = re.compile(
    r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$',
    re.IGNORECASE
)

def _is_uuid(name):
    # type: (str) -> bool
    """ファイル名（拡張子なし）がUUID形式かどうかを返す。"""
    stem = Path(name).stem
    return bool(_UUID_RE.match(stem))


# ===========================================================================
# 事前クリーンアップ
# ===========================================================================
def cleanup_downloads():
    # type: () -> None
    """
    downloads/ フォルダ内の全CSVファイルを削除する。
    """
    csv_files = list(DOWNLOAD_DIR.glob("*.csv"))
    if not csv_files:
        print("  downloads/ フォルダにCSVはありません。")
        return
    for f in csv_files:
        f.unlink()
        print("  削除: {}".format(f.name))
    print("  合計 {} 個のCSVを削除しました。".format(len(csv_files)))


# ===========================================================================
# Step 1: PlaywrightでCSVダウンロード（複数ファイル対応）
# ===========================================================================
def _file_hash(path):
    # type: (Path) -> str
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _save_download(download, index=0):
    # type: (...) -> Optional[Path]
    """
    download オブジェクトを downloads/ に保存し、パスを返す。
    ファイル名がUUID形式・拡張子なし・空の場合は連番CSVファイル名を使用する。
    """
    today = datetime.now().strftime("%Y%m%d")
    suggested = download.suggested_filename or ""

    # ファイル名がUUID形式、または拡張子が .csv でない場合は連番名に置き換える
    if not suggested or _is_uuid(suggested) or not suggested.lower().endswith(".csv"):
        filename = "{}{}_{}.csv".format(FILE_PREFIX, today, index)
    else:
        filename = suggested

    dest = DOWNLOAD_DIR / filename
    tmp  = DOWNLOAD_DIR / ("_tmp_" + filename)
    download.save_as(str(tmp))

    if dest.exists():
        if _file_hash(tmp) == _file_hash(dest):
            print("  [{}] 既に同じ内容のファイルが存在します: {} (スキップ)".format(index, dest.name))
            tmp.unlink()
            return dest
        else:
            ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
            dest = DOWNLOAD_DIR / "{}_{}_{}.csv".format(FILE_PREFIX, today, ts)

    tmp.rename(dest)
    print("  [{}] 保存しました: {}".format(index, dest.name))
    return dest


def download_csv():
    # type: () -> List[Path]
    """
    CAA届出情報DBページにアクセスし、
    「前日までの全届出の全項目出力(CSV出力)」ボタンを押して
    CSVファイルをダウンロードする。
    context.on("download", ...) イベントハンドラで全ダウンロードをキャプチャする。
    保存したファイルのパスのリストを返す。
    """
    print("[1/3] ページにアクセスしてCSVをダウンロードします...")
    print("  URL: {}".format(TARGET_URL))

    saved_paths = []   # type: List[Path]
    pending_downloads = []  # type: list  # Downloadオブジェクトのキュー

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)

        # --- コンテキストレベルのダウンロードハンドラ ---
        # メインページ・ポップアップを問わず全ダウンロードを捕捉
        def on_download(download):
            # type: (Download) -> None
            pending_downloads.append(download)

        context.on("download", on_download)

        page = context.new_page()
        page.goto(TARGET_URL, timeout=60_000)
        page.wait_for_load_state("networkidle", timeout=30_000)

        btn = page.get_by_text(BUTTON_TEXT, exact=False).first
        if not btn:
            print("ERROR: ボタン '{}' が見つかりません。".format(BUTTON_TEXT))
            browser.close()
            return []

        print("  ボタン発見: '{}' → クリック".format(btn.inner_text().strip()))
        btn.click()
        print("  ダウンロード完了を待っています... (最大{}秒)".format(DOWNLOAD_TIMEOUT))

        # ダウンロードが来るまで最大 DOWNLOAD_TIMEOUT 秒待機
        # 最後のダウンロードから NEXT_DOWNLOAD_WAIT 秒経過したら終了
        deadline = time.time() + DOWNLOAD_TIMEOUT
        last_count = 0
        last_new_at = time.time()

        while time.time() < deadline:
            time.sleep(1)
            current_count = len(pending_downloads)
            if current_count > last_count:
                last_count = current_count
                last_new_at = time.time()
                print("  ダウンロード検知: 累計 {} 件".format(current_count))
            elif last_count > 0:
                # 1件以上受信済みで NEXT_DOWNLOAD_WAIT 秒以上新着なし → 終了
                if time.time() - last_new_at >= NEXT_DOWNLOAD_WAIT:
                    print("  {}秒間新着なし → ダウンロード完了と判断".format(NEXT_DOWNLOAD_WAIT))
                    break

        if not pending_downloads:
            print("  タイムアウト: ダウンロードが検知されませんでした。")
            browser.close()
            return []

        # 全ダウンロードを保存
        for i, dl in enumerate(pending_downloads):
            path = _save_download(dl, index=i + 1)
            if path:
                saved_paths.append(path)

        browser.close()

    print("  合計 {} ファイルをダウンロードしました。".format(len(saved_paths)))
    return saved_paths


# ===========================================================================
# Step 2: 全CSVを統合
# ===========================================================================
def load_all_csvs():
    # type: () -> pd.DataFrame
    """
    downloads/ フォルダ内の全CSVを読み込み、統合した DataFrame を返す。
    1行目を見出し行として扱い、2行目以降を結合する。
    重複行（全列一致）は除去する。
    """
    csv_files = sorted(DOWNLOAD_DIR.glob("*.csv"))
    if not csv_files:
        print("ERROR: downloads/ フォルダにCSVファイルが見つかりません。")
        sys.exit(1)

    print("[2/3] {} 個のCSVを統合します...".format(len(csv_files)))
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
            print("  読み込み: {}  ({:,} 件)".format(f.name, len(df)))
        except UnicodeDecodeError:
            # UTF-8 フォールバック
            try:
                df = pd.read_csv(
                    f,
                    encoding="utf-8-sig",
                    dtype=str,
                    header=0,
                    on_bad_lines="skip",
                )
                frames.append(df)
                print("  読み込み(UTF-8): {}  ({:,} 件)".format(f.name, len(df)))
            except Exception as e:
                print("  警告: {} の読み込みをスキップしました ({})".format(f.name, e))
        except Exception as e:
            print("  警告: {} の読み込みをスキップしました ({})".format(f.name, e))

    if not frames:
        print("ERROR: 読み込めるCSVがありませんでした。")
        sys.exit(1)

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.fillna("")
    # 完全重複を除去
    combined = combined.drop_duplicates()
    print("  統合後: {:,} 件 / {} 列".format(len(combined), len(combined.columns)))
    return combined


# ===========================================================================
# Step 3: マスタファイルの作成 or 更新
# ===========================================================================
def update_master(new_df):
    # type: (pd.DataFrame) -> None
    """
    初回 → master.xlsx を新規作成
    2回目以降:
      - 届出番号が新規 → 追加
      - 届出番号が一致するが他列が異なる → 上書き
    """
    if not MASTER_PATH.exists():
        # 初回
        print("[3/3] master.xlsx を新規作成します: {}".format(MASTER_PATH))
        new_df.to_excel(str(MASTER_PATH), index=False, engine="openpyxl")
        print("  完了: {:,} 件".format(len(new_df)))
        return

    # 2回目以降
    print("[3/3] master.xlsx と差分を確認して更新します...")
    try:
        master_df = pd.read_excel(str(MASTER_PATH), dtype=str, header=0)
        master_df = master_df.fillna("")
    except Exception as e:
        print("ERROR: master.xlsx の読み込みに失敗しました: {}".format(e))
        sys.exit(1)

    # 見出し列をチェック
    if KEY_COL not in master_df.columns:
        print("WARNING: master.xlsx に '{}' 列が見つかりません。".format(KEY_COL))
        print("  全データを上書きします。")
        new_df.to_excel(str(MASTER_PATH), index=False, engine="openpyxl")
        print("  上書き完了: {:,} 件".format(len(new_df)))
        return

    if KEY_COL not in new_df.columns:
        print("WARNING: 新しいCSVに '{}' 列が見つかりません。".format(KEY_COL))
        print("  全データを上書きします。")
        new_df.to_excel(str(MASTER_PATH), index=False, engine="openpyxl")
        print("  上書き完了: {:,} 件".format(len(new_df)))
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

    print("  新規追加: {:,} 件".format(added))
    print("  上書き更新: {:,} 件".format(updated))
    print("  master.xlsx 更新完了: {:,} 件".format(len(result)))


# ===========================================================================
# メイン
# ===========================================================================
def main():
    print("=" * 60)
    print("機能性表示食品届出情報 CSV取得・マスタ更新スクリプト")
    print("  実行日時: {}".format(datetime.now().strftime("%Y年%m月%d日 %H:%M:%S")))
    print("=" * 60)

    # Step 0: downloads/ 内の旧CSVを削除
    print("[0/3] downloads/ 内の旧CSVをクリアします...")
    cleanup_downloads()

    # Step 1: CSVダウンロード（複数ファイル対応）
    saved_list = download_csv()
    if not saved_list:
        print("CSV のダウンロードに失敗しました。処理を中断します。")
        sys.exit(1)

    # Step 2: 全CSV統合
    combined = load_all_csvs()

    # Step 3: マスタ更新
    update_master(combined)

    print()
    print("すべての処理が完了しました。")
    print("  マスタファイル: {}".format(MASTER_PATH))


if __name__ == "__main__":
    main()
