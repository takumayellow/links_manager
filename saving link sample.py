#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import time
import subprocess
import argparse
from pywinauto import Desktop
from pywinauto.keyboard import send_keys

# ———— 引数処理 ————
parser = argparse.ArgumentParser(
    description="メモ操作: (o) リンクを開く / (s) Chromeウィンドウ選択→タブURLコピー＆削除＆メモ先頭追加"
)
parser.add_argument(
    "-m", "--memo-path",
    default=r"C:\Users\takum\Downloads\links.txt",
    help="メモファイルのフルパス（Downloads内の無難な名前）"
)
parser.add_argument(
    "-c", "--chrome-path",
    default=r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    help="Chrome実行ファイルのパス"
)
args = parser.parse_args()
MEMO_PATH = args.memo_path
CHROME_EXE = args.chrome_path


def open_and_remove_first_n():
    """メモの先頭から指定数リンクを通常モードのChromeで開き、行を削除"""
    if not os.path.exists(MEMO_PATH):
        print(f"[Error] メモが見つかりません: {MEMO_PATH}")
        sys.exit(1)
    with open(MEMO_PATH, 'r', encoding='utf-8') as f:
        lines = [l.strip() for l in f if l.strip()]
    if not lines:
        print("[Info] メモにリンクがありません。")
        return
    # 開く件数を入力
    try:
        n = int(input(f"何件開きますか？（最大{len(lines)}件） > ").strip())
    except Exception:
        print("[Error] 数字を入力してください。")
        return
    if n <= 0:
        print("[Info] 0件選択されたため何も行いません。")
        return
    to_open = lines[:n]
    rest = lines[n:]
    # 複数リンクを順番に開く
    for url in to_open:
        print(f"[Open] {url}")
        try:
            subprocess.Popen([CHROME_EXE, url])
        except Exception as e:
            print(f"[Warn] Chrome起動失敗: {e}")
        time.sleep(1)
    # 残りをファイルへ
    with open(MEMO_PATH, 'w', encoding='utf-8') as f:
        for l in rest:
            f.write(l + "\n")
    print(f"[Done] {len(to_open)} 件を開いてメモから削除しました。残り {len(rest)} 件です。")


def select_window_and_clear_tabs():
    """Chromeウィンドウを選択し、タブからURLをコピー、タブを削除、URLをメモ先頭に追加"""
    desktop = Desktop(backend='uia')
    wins = desktop.windows(title_re='.*Chrome.*', top_level_only=True)
    if not wins:
        print("[Warn] Chromeウィンドウが見つかりませんでした。")
        return
    print("=== Chromeウィンドウ一覧 ===")
    for idx, w in enumerate(wins):
        print(f"[{idx}] {w.window_text()}")
    print("==========================")
    try:
        sel = int(input("ウィンドウ番号を選択 > ").strip())
        win = wins[sel]
    except Exception:
        print("[Error] 無効な番号です。終了します。")
        return
    spec = Desktop(backend='uia').window(handle=win.handle)
    spec.set_focus()

    collected_urls = []
    while True:
        try:
            send_keys('^l')
            time.sleep(0.1)
            send_keys('^c')
            time.sleep(0.1)
            clip = subprocess.check_output('powershell Get-Clipboard', shell=True)
            url = clip.decode('utf-8').strip()
        except Exception:
            break
        if not url.startswith('http'):
            break
        collected_urls.append(url)
        send_keys('^w')
        time.sleep(0.2)
        try:
            spec.set_focus()
        except Exception:
            break

    if not collected_urls:
        print("[Info] 追加するタブはありませんでした。")
        return
    old_content = ''
    if os.path.exists(MEMO_PATH):
        with open(MEMO_PATH, 'r', encoding='utf-8') as f:
            old_content = f.read().rstrip("\n")
    new_content = "\n".join(collected_urls)
    final = new_content + ("\n" + old_content if old_content else '')
    with open(MEMO_PATH, 'w', encoding='utf-8') as f:
        f.write(final)
    print(f"[Done] {len(collected_urls)} 件のタブを削除しメモ先頭に追加しました。")


if __name__ == '__main__':
    print("選択: (o)リンクを開く / (s)ウィンドウ選択→タブ削除＆記録")
    c = input('[o/s]> ').strip().lower()
    if c == 'o':
        open_and_remove_first_n()
    elif c == 's':
        select_window_and_clear_tabs()
    else:
        print("[Error] 無効です。終了します。")
        sys.exit(1)
