# セットアップガイド

このガイドでは、メール自動化システムのセットアップ方法を詳しく説明します。

## 前提条件

- Python 3.11以降がインストールされている
- OpenAI APIキーを取得済み
- Outlookアカウントを持っている

## ステップ1: ファイルの準備

1. すべてのファイルをダウンロード
2. 任意のディレクトリに展開

## ステップ2: 依存パッケージのインストール

### Windows（PowerShell）

```powershell
# 仮想環境の作成
python -m venv venv

# 仮想環境のアクティベート
.\venv\Scripts\Activate.ps1

# 依存パッケージのインストール
pip install -r requirements.txt
```

### macOS/Linux

```bash
# セットアップスクリプトを実行
chmod +x setup.sh
./setup.sh
```

## ステップ3: 環境変数の設定

1. `env.example`を`.env`にコピー

```bash
cp env.example .env
```

2. `.env`ファイルを編集

### 最小構成（Windows + Outlook COM）

```bash
# メール設定
EMAIL_ADDRESS=your-email@company.com
MAIL_PROCESSOR_TYPE=outlook_com

# OpenAI API
OPENAI_API_KEY=sk-...

# Obsidian
OBSIDIAN_VAULT_PATH=G:\path\to\vault

# 処理モード
MARK_AS_READ_AFTER_PROCESS=0

# 個人情報
MY_NAME=山田
MY_NAME_INTERNAL=山田（太郎）
MY_COMPANY_NAME=株式会社サンプル

# ドメイン設定
INTERNAL_DOMAINS=@company.com
```

### 推奨設定（フル機能）

```bash
# メール設定
EMAIL_ADDRESS=your-email@company.com
MAIL_PROCESSOR_TYPE=outlook_com

# OpenAI API
OPENAI_API_KEY=sk-...

# Obsidian
OBSIDIAN_VAULT_PATH=G:\path\to\vault

# 処理モード
MARK_AS_READ_AFTER_PROCESS=0

# 個人情報
MY_NAME=山田
MY_NAME_INTERNAL=山田（太郎）
MY_COMPANY_NAME=株式会社サンプル

# ドメイン設定
INTERNAL_DOMAINS=@company.com
GROUP_EMAIL_ADDRESSES=team@company.com,project@company.com

# 除外設定
EXCLUDE_SENDER_DOMAINS=@teams.mail.microsoft
EXCLUDE_AND_READ_DOMAINS=newsletter@example.com,@sales.example.com

# プロジェクトキーワード
PROJECT_KEYWORDS=プロジェクトA,プロジェクトB

# Slack通知（オプション）
SLACK_BOT_TOKEN=xoxb-...
SLACK_CHANNEL=#monitoring
```

## ステップ4: テンプレートのカスタマイズ

`templates/` フォルダ内のテンプレートを編集します。

### 社内用テンプレート

`templates/internal_template.txt`を編集：

```
{sender_name}様

お疲れさまです、{my_name}です。

{body}

よろしくお願いいたします。
```

### 社外用テンプレート

`templates/external_template.txt`を編集：

```
{sender_company}　{sender_name}様

お世話になっております。
{my_company}の{my_name}です。

{body}

以上、よろしくお願いいたします。
```

## ステップ5: 初回実行

### Windows

```powershell
# 仮想環境をアクティベート（まだの場合）
.\venv\Scripts\Activate.ps1

# 実行
python main.py
```

### macOS/Linux

```bash
./run.sh
```

## ステップ6: ログの確認

実行後、ログファイルを確認します：

```bash
# 最新のログ
cat logs/mail_automation_2025-11-08.log

# リアルタイムで監視（macOS/Linux）
tail -f logs/mail_automation_*.log
```

正常に動作していれば、以下のようなログが表示されます：

```
INFO - Outlook COMオブジェクトを使用します
INFO - [設定確認] MARK_AS_READ_AFTER_PROCESS = 0
INFO - 未読メール数: 5
INFO - [判定] To に自分のアドレスあり → 要約+返信下書き → 未読維持
INFO - 📧 [パターン1] 要約＋返信下書き＋Obsidian保存を実行します
INFO - ✉️ [未読に戻しました] メール件名
```

## ステップ7: 自動実行の設定

### Windows タスクスケジューラ

1. タスクスケジューラを開く
2. 「基本タスクの作成」
3. トリガー: 10分ごと
4. 操作: プログラムの起動
   - プログラム: `C:\path\to\venv\Scripts\python.exe`
   - 引数: `main.py`
   - 開始: プロジェクトディレクトリ

### macOS LaunchAgent

```bash
chmod +x install_launchagent.sh
./install_launchagent.sh
```

## トラブルシューティング

### 問題1: メールが取得できない

**解決策:**
1. `.env`ファイルの`EMAIL_ADDRESS`が正しいか確認
2. Outlookが起動しているか確認（Outlook COM使用時）
3. ログファイルでエラーを確認

### 問題2: AI処理でエラー

**解決策:**
1. OpenAI APIキーが正しいか確認
2. OpenAI APIのクォータを確認
3. ログファイルで詳細なエラーを確認

### 問題3: テンプレートが適用されない

**解決策:**
1. テンプレートファイルのパスが正しいか確認
2. テンプレート内の変数名が正しいか確認（`{sender_name}`など）
3. `.env`の`INTERNAL_DOMAINS`が設定されているか確認

### 問題4: 想定と違うメール処理

**解決策:**
1. ログレベルを`DEBUG`に設定
2. 判定ログを確認
3. 設定を調整

```bash
LOG_LEVEL=DEBUG
```

## 設定のヒント

### ヒント1: まずはシンプルに

最初は全メールを未読のまま要約する設定がおすすめ：

```bash
MARK_AS_READ_AFTER_PROCESS=1
```

### ヒント2: 徐々に細かく

慣れてきたらルールベースに切り替え：

```bash
MARK_AS_READ_AFTER_PROCESS=0
MY_NAME=山田
GROUP_EMAIL_ADDRESSES=team@company.com
```

### ヒント3: 除外リストを活用

不要なメールは除外リストに追加：

```bash
EXCLUDE_AND_READ_DOMAINS=newsletter@example.com
```

## 次のステップ

1. 数日間動作を観察
2. ログを確認して調整
3. 除外リストを最適化
4. テンプレートをカスタマイズ
5. プロジェクトキーワードを追加

## サポート

質問や問題がある場合は、ログファイルを確認してください：

```bash
logs/mail_automation_YYYY-MM-DD.log
```

詳細なデバッグ情報が必要な場合は：

```bash
LOG_LEVEL=DEBUG
```

を設定して再実行してください。

