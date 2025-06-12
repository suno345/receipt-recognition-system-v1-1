#!/bin/bash

# 証憑画像認識システム v1.3統合版 - 複数商品対応版
# macOS用起動スクリプト

# スクリプトのディレクトリに移動
cd "$(dirname "$0")"

# 実行環境の確認
echo "🔍 実行環境を確認中..."

# Python3の確認
if command -v python3 &> /dev/null; then
    echo "✅ Python3が見つかりました"
    PYTHON_CMD="python3"
elif command -v python &> /dev/null; then
    # pythonコマンドがPython3かどうか確認
    if python --version 2>&1 | grep -q "Python 3"; then
        echo "✅ Pythonが見つかりました"
        PYTHON_CMD="python"
    else
        echo "❌ Python3が見つかりません"
        echo "Python3をインストールしてください"
        read -p "Enterで終了..."
        exit 1
    fi
else
    echo "❌ Pythonが見つかりません"
    echo "Python3をインストールしてください"
    read -p "Enterで終了..."
    exit 1
fi

# 必要なライブラリのインストール確認
echo ""
echo "📦 必要なライブラリを確認中..."

# pip3のインストール確認
if ! $PYTHON_CMD -m pip --version &> /dev/null; then
    echo "⚠️ pipがインストールされていません"
    echo "pipをインストール中..."
    $PYTHON_CMD -m ensurepip --default-pip
fi

# 必要なライブラリをインストール
REQUIRED_PACKAGES="openai openpyxl pillow python-dotenv"

for package in $REQUIRED_PACKAGES; do
    if ! $PYTHON_CMD -c "import $package" &> /dev/null 2>&1; then
        echo "📥 $package をインストール中..."
        $PYTHON_CMD -m pip install $package
    fi
done

echo ""
echo "✅ 環境準備が完了しました"
echo ""
echo "=" * 60
echo "🚀 証憑画像認識システム（複数商品対応版）を起動します..."
echo "=" * 60
echo ""

# メインプログラムを実行
$PYTHON_CMD "システムファイル/main.py"