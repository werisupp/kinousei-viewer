#!/bin/bash
# Mac用：このファイルをダブルクリックするとPythonが実行されブラウザが開きます
# 初回のみターミナルで以下を実行してください:
#   chmod +x 更新してブラウザで開く.command

# このスクリプトのあるフォルダに移動
cd "$(dirname "$0")"

echo "===================================="
echo " 機能性表示食品 ビューワー 起動中"
echo "===================================="
echo ""

# pandasがインストールされているか確認
python3 -c "import pandas" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "pandas をインストールしています..."
    pip3 install pandas openpyxl
    if [ $? -ne 0 ]; then
        echo ""
        echo "エラー：ライブラリのインストールに失敗しました。"
        echo "ターミナルで以下を実行してください:"
        echo "  pip3 install pandas openpyxl"
        echo ""
        read -p "Enterで閉じる..."
        exit 1
    fi
fi

# スクリプト実行
python3 build_viewer.py

if [ $? -ne 0 ]; then
    echo ""
    echo "エラーが発生しました。上記のメッセージを確認してください。"
    read -p "Enterで閉じる..."
fi
