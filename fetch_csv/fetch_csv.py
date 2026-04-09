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
import sys
import hashlib
import time
from datetime import datetime
from pathlib import Path
from typing import Optional, List

# ライブラリ確認 ----------------------------------------------------------------
def _require(pkg, install_hint):
    try:
        __import__(pkg)
    except ImportError:
        print("ERROR: {} がありません。".format(pkg))
        print("  {}".format(install_hint))
        sys.exit(1)

_require("playwright", "pip install playwright && python3 -m playwright install chromium")
_require("requests",   "pip install requests")
_require("pandas",     "pip install pandas openpyxl")
_require("openpyxl",   "pip install openpyxl")

import requests  # noqa: E402
import pandas as pd  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

# パス設定 ---------------------------------------------------------------------
SCRIPT_DIR   = Path(__file__).resolve().parent
DOWNLOAD_DIR = SCRIPT_DIR / "downloads"
MASTER_PATH  = SCRIPT_DIR / "master.xlsx"
TARGET_URL   = "https://www.fld.caa.go.jp/caaks/s/cssc01/"
BUTTON_TEXT  = "前日までの全届出の全項目出力"
KEY_COL      = "届出番号"

# ボタンクリック後、ネットワークリクエストを待つ最大秒数
INTERCEPT_TIMEOUT = 120  # 秒

# 最後のリクエスト検知から追加を待つ秒数
NEXT_REQUEST_WAIT = 30  # 秒

DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

# ===========================================================================
# 事前クリーンアップ
# ===========================================================================
def cleanup_downloads():
    # type: () -> None
    csv_files = list(DOWNLOAD_DIR.glob("*.csv"))
    if not csv_files:
        print("  downloads/ フォルダにCSVはありません。")
        return
    for f in csv_files:
        f.unlink()
        print("  削除: {}".format(f.name))
    print("  合計 {} 個のCSVを削除しました。".format(len(csv_files)))


# ===========================================================================
# Step 1: ネットワークインターセプトでURL取得 → requestsで保存
# ===========================================================================
def download_csv():
    # type: () -> List[Path]
    """
    Playwrightでボタンクリック後に発生するネットワークリクエストをインターセプトし、
    CSV配信URL・Cookieを取得。その後 requests で実際にファイルを保存する。
    """
    print("[1/3] ページにアクセスしてCSV配信URLを取得します...")
    print("  URL: {}".format(TARGET_URL))

    csv_urls = []     # type: List[str]
    cookies_dict = {} # type: dict
    headers_dict = {} # type: dict

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # ネットワークリクエストをモニタリング
        # CSV配信のリクエスト（大容量・拡張子なし）を捕捉する
        def on_request(request):
            url = request.url
            # メインページやJS/CSSを除く、API系リクエストを捕捉
            if (
                "cssc01" not in url  # メインページは除外
                and url not in csv_urls
                and not url.endswith((".js", ".css", ".png", ".jpg", ".ico", ".woff", ".woff2"))
                and "fld.caa.go.jp" in url
            ):
                csv_urls.append(url)
                print("  リクエスト検知: {}".format(url))

        page.on("request", on_request)

        page.goto(TARGET_URL, timeout=60_000)
        page.wait_for_load_state("networkidle", timeout=30_000)

        btn = page.get_by_text(BUTTON_TEXT, exact=False).first
        if not btn:
            print("ERROR: ボタン '{}' が見つかりません。".format(BUTTON_TEXT))
            browser.close()
            return []

        print("  ボタン発見: '{}' → クリック".format(btn.inner_text().strip()))

        # ボタンクリック前のURLリストをリセット
        csv_urls.clear()
        btn.click()

        print("  リクエスト検知中... (最大{}秒)".format(INTERCEPT_TIMEOUT))

        deadline = time.time() + INTERCEPT_TIMEOUT
        last_count = 0
        last_new_at = time.time()

        while time.time() < deadline:
            time.sleep(1)
            current_count = len(csv_urls)
            if current_count > last_count:
                last_count = current_count
                last_new_at = time.time()
                print("  URL検知: 累計 {} 件".format(current_count))
            elif last_count > 0:
                if time.time() - last_new_at >= NEXT_REQUEST_WAIT:
                    print("  {}秒間新着なし → 取得完了と判断".format(NEXT_REQUEST_WAIT))
                    break

        # Cookieをrequests用に取得
        browser_cookies = context.cookies()
        for c in browser_cookies:
            cookies_dict[c["name"]] = c["value"]

        # User-Agentを取得
        ua = page.evaluate("navigator.userAgent")
        headers_dict = {
            "User-Agent": ua,
            "Referer": TARGET_URL,
        }

        browser.close()

    if not csv_urls:
        print("  URLが検知されませんでした。")
        return []

    print("  検知したURL: {} 件 → requestsで保存開始".format(len(csv_urls)))

    saved_paths = []  # type: List[Path]
    date_str = datetime.now().strftime("%Y%m%d")

    for i, url in enumerate(csv_urls):
        label = "{}/{}".format(i + 1, len(csv_urls))
        dest = DOWNLOAD_DIR / "kinousei_{}_{:02d}.csv".format(date_str, i + 1)
        print("  [{}] ダウンロード中: {}".format(label, url))
        try:
            resp = requests.get(url, cookies=cookies_dict, headers=headers_dict, timeout=120, stream=True)
            resp.raise_for_status()
            with open(dest, "wb") as f:
                for chunk in resp.iter_content(chunk_size=65536):
                    if chunk:
                        f.write(chunk)
            size_mb = dest.stat().st_size / 1024 / 1024
            print("  [{}] 保存完了: {} ({:.1f} MB)".format(label, dest.name, size_mb))
            saved_paths.append(dest)
        except Exception as e:
            print("  [{}] エラー: {}".format(label, e))

    print("  合計 {} ファイルを保存しました。".format(len(saved_paths)))
    return saved_paths


# ===========================================================================
# Step 2: 全CSVを統合
# ===========================================================================
def load_all_csvs():
    # type: () -> pd.DataFrame
    csv_files = sorted(DOWNLOAD_DIR.glob("*.csv"))
    if not csv_files:
        print("ERROR: downloads/ フォルダにCSVファイルが見つかりません。")
        sys.exit(1)

    print("[2/3] {} 個のCSVを統合します...".format(len(csv_files)))
    frames = []
    for f in csv_files:
        try:
            df = pd.read_csv(f, encoding="cp932", dtype=str, header=0, on_bad_lines="skip")
            frames.append(df)
            print("  読み込み: {}  ({:,} 件)".format(f.name, len(df)))
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(f, encoding="utf-8-sig", dtype=str, header=0, on_bad_lines="skip")
                frames.append(df)
                print("  読み込み(UTF-8): {}  ({:,} 件)".format(f.name, len(df)))
            except Exception as e:
                print("  警告: {} の読み込みをスキップ ({})".format(f.name, e))
        except Exception as e:
            print("  警告: {} の読み込みをスキップ ({})".format(f.name, e))

    if not frames:
        print("ERROR: 読み込めるCSVがありませんでした。")
        sys.exit(1)

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.fillna("").drop_duplicates()
    print("  統合後: {:,} 件 / {} 列".format(len(combined), len(combined.columns)))
    return combined


# ===========================================================================
# Step 3: マスタファイルの作成 or 更新
# ===========================================================================
def update_master(new_df):
    # type: (pd.DataFrame) -> None
    if not MASTER_PATH.exists():
        print("[3/3] master.xlsx を新規作成します: {}".format(MASTER_PATH))
        new_df.to_excel(str(MASTER_PATH), index=False, engine="openpyxl")
        print("  完了: {:,} 件".format(len(new_df)))
        return

    print("[3/3] master.xlsx と差分を確認して更新します...")
    try:
        master_df = pd.read_excel(str(MASTER_PATH), dtype=str, header=0).fillna("")
    except Exception as e:
        print("ERROR: master.xlsx の読み込みに失敗: {}".format(e))
        sys.exit(1)

    if KEY_COL not in master_df.columns or KEY_COL not in new_df.columns:
        print("WARNING: '{}' 列が見つかりません。全データを上書きします。".format(KEY_COL))
        new_df.to_excel(str(MASTER_PATH), index=False, engine="openpyxl")
        print("  上書き完了: {:,} 件".format(len(new_df)))
        return

    all_cols = list(dict.fromkeys(list(master_df.columns) + list(new_df.columns)))
    for col in all_cols:
        if col not in master_df.columns:
            master_df[col] = ""
        if col not in new_df.columns:
            new_df[col] = ""
    master_df = master_df[all_cols]
    new_df    = new_df[all_cols]

    master_indexed = master_df.set_index(KEY_COL)
    new_indexed    = new_df.set_index(KEY_COL)

    added = updated = 0
    for key, new_row in new_indexed.iterrows():
        if key not in master_indexed.index:
            master_indexed = pd.concat([master_indexed, new_row.to_frame().T])
            added += 1
        else:
            existing = master_indexed.loc[key]
            if isinstance(existing, pd.DataFrame):
                existing = existing.iloc[0]
            cols_to_check = [c for c in all_cols if c != KEY_COL]
            if not existing[cols_to_check].equals(new_row[cols_to_check]):
                master_indexed.loc[key] = new_row
                updated += 1

    result = master_indexed.reset_index()
    result.to_excel(str(MASTER_PATH), index=False, engine="openpyxl")
    print("  新規追加: {:,} 件 / 上書き更新: {:,} 件".format(added, updated))
    print("  master.xlsx 更新完了: {:,} 件".format(len(result)))


# ===========================================================================
# メイン
# ===========================================================================
def main():
    print("=" * 60)
    print("機能性表示食品届出情報 CSV取得・マスタ更新スクリプト")
    print("  実行日時: {}".format(datetime.now().strftime("%Y年%m月%d日 %H:%M:%S")))
    print("=" * 60)

    print("[0/3] downloads/ 内の旧CSVをクリアします...")
    cleanup_downloads()

    saved_list = download_csv()
    if not saved_list:
        print("CSV のダウンロードに失敗しました。処理を中断します。")
        sys.exit(1)

    combined = load_all_csvs()
    update_master(combined)

    print()
    print("すべての処理が完了しました。")
    print("  マスタファイル: {}".format(MASTER_PATH))


if __name__ == "__main__":
    main()
