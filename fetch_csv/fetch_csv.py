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
import threading
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
from playwright.sync_api import sync_playwright, Page, BrowserContext  # noqa: E402

# パス設定 ---------------------------------------------------------------------
SCRIPT_DIR   = Path(__file__).resolve().parent          # fetch_csv/
DOWNLOAD_DIR = SCRIPT_DIR / "downloads"                 # fetch_csv/downloads/
MASTER_PATH  = SCRIPT_DIR / "master.xlsx"               # fetch_csv/master.xlsx
TARGET_URL   = "https://www.fld.caa.go.jp/caaks/s/cssc01/"
BUTTON_TEXT  = "前日までの全届出の全項目出力"
KEY_COL      = "届出番号"                                # マスタ更新のキー列

# ダウンロード完了を待つ最大秒数（全ファイル合計）
DOWNLOAD_TIMEOUT = 300_000  # 300秒

DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

# ===========================================================================
# Step 1: PlaywrightでCSVダウンロード（複数ファイル対応）
# ===========================================================================
def _save_download(download, label=""):
    # type: (...) -> Optional[Path]
    """download オブジェクトを downloads/ に保存し、パスを返す。"""
    suggested = download.suggested_filename or "kinousei_{}.csv".format(
        datetime.now().strftime("%Y%m%d_%H%M%S")
    )
    dest = DOWNLOAD_DIR / suggested
    tmp  = DOWNLOAD_DIR / ("_tmp_" + suggested)
    download.save_as(str(tmp))

    if dest.exists():
        if _file_hash(tmp) == _file_hash(dest):
            print("  [{}] 既に同じ内容のファイルが存在します: {} (スキップ)".format(label, dest.name))
            tmp.unlink()
            return dest
        else:
            ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
            dest = DOWNLOAD_DIR / "{}_{{}}{{}}" .format(dest.stem).format(ts, dest.suffix)

    tmp.rename(dest)
    print("  [{}] 保存しました: {}".format(label, dest.name))
    return dest


def _handle_page_downloads(page, saved_paths, lock, label="popup"):
    # type: (Page, list, threading.Lock, str) -> None
    """
    ポップアップページで発生するダウンロードをすべて収集する（複数ファイル対応）。
    ダウンロードが一定時間発生しなくなったら終了する。
    """
    import time
    deadline = time.time() + (DOWNLOAD_TIMEOUT / 1000)
    while time.time() < deadline:
        try:
            remaining_ms = int((deadline - time.time()) * 1000)
            if remaining_ms <= 0:
                break
            # 次のダウンロードを最大30秒待つ（30秒以上来なければ終了）
            wait_ms = min(remaining_ms, 30_000)
            with page.expect_download(timeout=wait_ms) as dl_info:
                pass
            path = _save_download(dl_info.value, label=label)
            if path:
                with lock:
                    saved_paths.append(path)
        except Exception:
            # タイムアウト → これ以上ダウンロードなしと判断して終了
            break


def download_csv():
    # type: () -> List[Path]
    """
    CAA届出情報DBページにアクセスし、
    「前日までの全届出の全項目出力(CSV出力)」ボタンを押して
    CSVファイルをダウンロードする。
    ・メインページのダウンロード
    ・ポップアップ（新しいウィンドウ）経由のダウンロード
    の両方を捕捉する。
    保存したファイルのパスのリストを返す。
    """
    print("[1/3] ページにアクセスしてCSVをダウンロードします...")
    print("  URL: {}".format(TARGET_URL))

    saved_paths = []   # type: List[Path]
    lock = threading.Lock()

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        context = browser.new_context(
            accept_downloads=True,
            # ポップアップをブロックしない（デフォルトは許可）
        )
        page = context.new_page()

        # ポップアップ（新しいページ）が開かれたときのハンドラを登録
        # ★ btn.click() より前に登録しておくことで取りこぼしを防ぐ
        popup_threads = []  # type: List[threading.Thread]

        def on_page(popup):
            # type: (Page) -> None
            popup.set_default_timeout(DOWNLOAD_TIMEOUT)
            t = threading.Thread(
                target=_handle_page_downloads,
                args=(popup, saved_paths, lock, "popup"),
                daemon=True,
            )
            popup_threads.append(t)
            t.start()

        context.on("page", on_page)

        page.goto(TARGET_URL, timeout=60_000)
        page.wait_for_load_state("networkidle", timeout=30_000)

        # ボタンを探す（テキストに「前日までの全届出の全項目出力」を含む要素）
        btn = page.get_by_text(BUTTON_TEXT, exact=False).first
        if not btn:
            print("ERROR: ボタン '{}' が見つかりません。".format(BUTTON_TEXT))
            browser.close()
            return []

        print("  ボタン発見: '{}' → クリック".format(btn.inner_text().strip()))

        # ★ メインページでダウンロードが発生する場合に備えて待受しつつ、
        #    1回だけクリックする。ポップアップのみの場合は10秒でタイムアウトし
        #    except に流れるが、その場合でも on_page ハンドラが捕捉している。
        try:
            with page.expect_download(timeout=10_000) as dl_main:
                btn.click()
            path = _save_download(dl_main.value, label="main")
            if path:
                with lock:
                    saved_paths.append(path)
        except Exception:
            # メインページではダウンロードが発生しない（ポップアップのみ）→ 正常
            # btn.click() は expect_download のコンテキスト内で1回だけ実行済み
            pass

        # ポップアップスレッドがすべて完了するまで待つ
        for t in popup_threads:
            t.join(timeout=DOWNLOAD_TIMEOUT / 1000)

        browser.close()

    print("  合計 {} ファイルをダウンロードしました。".format(len(saved_paths)))
    return saved_paths


def _file_hash(path):
    # type: (Path) -> str
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


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
