#!/bin/bash

# boothcsv.sh
# macOS/Linux用の簡易HTTPサーバー起動スクリプト
# - カレントディレクトリを配信
# - デフォルトページとしてboothcsv.htmlを返す

# デフォルト設定
PORT=8796
DEFAULT_PAGE="boothcsv.html"
OPEN_BROWSER=true

# コマンドライン引数の処理
while [[ $# -gt 0 ]]; do
  case $1 in
    -p|--port)
      PORT="$2"
      shift 2
      ;;
    --no-browser)
      OPEN_BROWSER=false
      shift
      ;;
    -h|--help)
      echo "使用方法: $0 [オプション]"
      echo "オプション:"
      echo "  -p, --port PORT     ポート番号を指定 (デフォルト: 8796)"
      echo "  --no-browser        ブラウザを自動起動しない"
      echo "  -h, --help          このヘルプを表示"
      exit 0
      ;;
    *)
      echo "不明なオプション: $1"
      echo "使用方法: $0 --help"
      exit 1
      ;;
  esac
done

# 現在のディレクトリにboothcsv.htmlがあるかチェック
if [[ ! -f "$DEFAULT_PAGE" ]]; then
  echo "エラー: $DEFAULT_PAGE が見つかりません"
  echo "boothcsv.htmlがあるディレクトリで実行してください"
  exit 1
fi

# Pythonの確認
if command -v python3 &> /dev/null; then
  PYTHON_CMD="python3"
elif command -v python &> /dev/null; then
  PYTHON_VERSION=$(python --version 2>&1)
  if [[ $PYTHON_VERSION == *"Python 3"* ]]; then
    PYTHON_CMD="python"
  else
    echo "エラー: Python 3が見つかりません"
    echo "Python 3をインストールするか、以下のコマンドを直接実行してください:"
    echo "python3 -m http.server $PORT"
    exit 1
  fi
else
  echo "エラー: Pythonが見つかりません"
  echo "Pythonをインストールするか、他のHTTPサーバーを使用してください"
  exit 1
fi

echo "🚀 BOOTHCSVサーバーを起動しています..."
echo "📁 配信ディレクトリ: $(pwd)"
echo "🌐 URL: http://localhost:$PORT"
echo "⏹️  停止するには Ctrl+C を押してください"
echo ""

# ブラウザを開く（バックグラウンドで）
if [[ "$OPEN_BROWSER" == true ]]; then
  # macOSの場合
  if [[ "$OSTYPE" == "darwin"* ]]; then
    open "http://localhost:$PORT" &
  # Linuxの場合
  elif command -v xdg-open &> /dev/null; then
    xdg-open "http://localhost:$PORT" &
  elif command -v gnome-open &> /dev/null; then
    gnome-open "http://localhost:$PORT" &
  else
    echo "ℹ️  ブラウザで http://localhost:$PORT を開いてください"
  fi
fi

# Ctrl+Cのシグナルハンドリング
trap 'echo -e "\n\n🛑 サーバーを停止しています..."; exit 0' INT

# HTTPサーバーを起動
$PYTHON_CMD -m http.server $PORT
