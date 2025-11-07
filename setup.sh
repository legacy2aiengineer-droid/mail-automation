#!/bin/bash

# メール自動化プロジェクト セットアップスクリプト
# macOS用

echo "==================================="
echo "メール自動化プロジェクト セットアップ"
echo "==================================="

# 1. Python仮想環境の作成
echo "Python仮想環境を作成中..."
python3 -m venv venv

# 2. 仮想環境をアクティベート
echo "仮想環境をアクティベート中..."
source venv/bin/activate

# 3. pipのアップグレード
echo "pipをアップグレード中..."
pip install --upgrade pip

# 4. 依存パッケージのインストール
echo "依存パッケージをインストール中..."
pip install -r requirements.txt

# 5. ディレクトリの作成
echo "必要なディレクトリを作成中..."
mkdir -p logs
mkdir -p credentials

# 6. .envファイルの確認
if [ ! -f .env ]; then
    echo ".envファイルが見つかりません。"
    echo ".env.exampleをコピーして .env を作成してください。"
    cp .env.example .env
    echo ".envファイルを作成しました。環境変数を設定してください。"
else
    echo ".envファイルが見つかりました。"
fi

# 7. 権限設定
echo "スクリプトに実行権限を付与中..."
chmod +x run.sh
chmod +x install_launchagent.sh

echo ""
echo "==================================="
echo "セットアップ完了！"
echo "==================================="
echo ""
echo "次の手順:"
echo "1. .envファイルを編集して環境変数を設定"
echo "2. credentials/google-credentials.json を配置"
echo "3. './run.sh' で手動テスト実行"
echo "4. './install_launchagent.sh' で自動実行を設定"
echo ""

