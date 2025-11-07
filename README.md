# メール自動化システム

Outlookメールの自動処理システムです。受信メールの要約、自動返信の下書き作成、監視メールの異常検知、TODOリストの自動抽出などを自動化します。

## 特徴

### 🎯 インテリジェントなメール分類
- **パターン1（要約+返信下書き）**: Toに自分のアドレスが含まれるメール
- **パターン2（要約のみ）**: CC/グループ宛で本文に名前が含まれるメール
- **パターン0（既読のみ）**: その他のメール
- **除外メール**: Teams通知やセールスメールを自動除外

### 📧 社内外対応テンプレート
- 社内メール: フレンドリーな文体
- 社外メール: フォーマルな文体
- テンプレートファイルで自由にカスタマイズ可能

### 🤖 AI機能
- メール内容の自動要約
- 文体学習による自然な返信文生成
- TODO自動抽出
- 異常メール検知

### 📝 Obsidian連携
- メール内容を自動保存
- 日次インデックス自動生成
- 送信者別インデックス
- タグとリンクで関連付け

## 機能

- **フローA**: メール受信→要約→下書き→添付保存→Obsidian保存
- **フローB**: 運用定期メールの異常検知→Slack通知
- **フローC**: 日次メールメトリクスレポート（8:00以降、1日1回）
- **フローD**: 取りこぼしメールの回収（5〜10分毎）
- **フローE**: TODOリストの自動抽出

## 必要な環境

- macOS 10.14以降 または Windows 10以降
- Python 3.11以降
- Office 365 アカウント（IMAP/SMTP有効）
- OpenAI API キー
- Google Drive API 認証情報（オプション）
- Slack Bot Token（オプション）

## セットアップ

### 1. リポジトリのクローン

```bash
git clone <repository-url>
cd mail-automation
```

### 2. セットアップスクリプトの実行

```bash
chmod +x setup.sh
./setup.sh
```

### 3. 環境変数の設定

`env.example`を`.env`にコピーして、環境変数を設定してください：

```bash
cp env.example .env
```

`.env`ファイルを編集：

```bash
# 基本設定
EMAIL_ADDRESS=your-email@example.com
OPENAI_API_KEY=your-openai-api-key
MAIL_PROCESSOR_TYPE=outlook_com  # Windows推奨

# 個人情報
MY_NAME=山田
MY_NAME_INTERNAL=山田（太郎）
MY_COMPANY_NAME=株式会社サンプル

# ドメイン設定
INTERNAL_DOMAINS=@example.com
GROUP_EMAIL_ADDRESSES=team@example.com

# Obsidian
OBSIDIAN_VAULT_PATH=G:\path\to\vault
```

詳細な設定項目は `env.example` を参照してください。

### 4. テンプレートのカスタマイズ

返信メールのテンプレートを編集できます：

- 社内用: `templates/internal_template.txt`
- 社外用: `templates/external_template.txt`

使用可能な変数：
- `{sender_name}`: 送信者の名前
- `{sender_company}`: 送信者の会社名
- `{my_name}`: 自分の名前
- `{my_company}`: 自分の会社名
- `{body}`: AIが生成する返信本文

### 5. Google Drive認証情報の配置（オプション）

`credentials/google-credentials.json` にGoogle Cloud Consoleから取得した認証情報を配置してください。

## 使用方法

### 手動実行

```bash
./run.sh
```

### 自動実行の設定（macOS）

```bash
chmod +x install_launchagent.sh
./install_launchagent.sh
```

これで10分毎に自動実行されます。

### Windows タスクスケジューラ

Windows環境では、タスクスケジューラを使用して定期実行を設定できます。

## ディレクトリ構造

```
mail-automation/
├── .env                    # 環境変数（要作成）
├── env.example             # 環境変数テンプレート
├── .gitignore              # Git除外設定
├── requirements.txt        # 依存パッケージ
├── setup.sh                # セットアップスクリプト
├── run.sh                  # 手動実行スクリプト
├── install_launchagent.sh  # LaunchAgentインストール
├── README.md               # このファイル
├── config.py               # 設定管理
├── main.py                 # メイン処理
├── mail_processor.py       # メール処理（IMAP/SMTP）
├── outlook_com_processor.py # メール処理（Outlook COM）
├── ews_mail_processor.py   # メール処理（EWS）
├── graph_mail_processor.py # メール処理（Graph API）
├── ai_processor.py         # AI処理
├── storage_manager.py      # 保存処理
├── utils.py                # ユーティリティ
├── templates/              # メールテンプレート
│   ├── internal_template.txt  # 社内用
│   └── external_template.txt  # 社外用
├── logs/                   # ログディレクトリ
└── credentials/            # 認証情報（要配置）
    └── google-credentials.json
```

## メール処理フロー

```
受信メール
  ↓
【除外判定（最優先）】
  ├─ セールスメール (@newsletter.example.com) → 🗑️ 既読
  ├─ Teams通知 (@teams.mail.microsoft) → 🚫 未読維持
  ↓
【通常判定】
  ├─ To: 自分宛 → 📧 要約+返信下書き（未読維持）
  ├─ CC: 自分宛 → 📬 要約のみ（未読維持）
  ├─ グループ宛+名前あり → 📬 要約のみ（未読維持）
  ├─ BCC+名前あり → 📬 要約のみ（未読維持）
  └─ それ以外 → ✓ 既読
```

## 主な機能

### 1. メール処理パターン

| パターン | 条件 | AI要約 | 返信生成 | 既読/未読 |
|---------|-----|-------|---------|---------|
| 除外（既読） | セールスメール等 | ❌ | ❌ | 既読 |
| 除外（未読） | Teams通知等 | ❌ | ❌ | 未読 |
| パターン1 | To: 自分宛 | ✅ | ✅ | 未読 |
| パターン2 | CC/グループ+名前 | ✅ | ❌ | 未読 |
| パターン0 | その他 | ❌ | ❌ | 既読 |

### 2. 社内外テンプレート

**社内用テンプレート** (`templates/internal_template.txt`):
```
{sender_name}様

お疲れさまです、{my_name}です。

{body}

よろしくお願いいたします。
```

**社外用テンプレート** (`templates/external_template.txt`):
```
{sender_company}　{sender_name}様

お世話になっております。
{my_company}の{my_name}です。

{body}

以上、よろしくお願いいたします。
```

### 3. フローC（日次レポート）

- 8時以降の初回実行時に1日1回実行
- 重複実行を自動防止
- 前日分のメール統計を生成

## トラブルシューティング

### 認証エラー

トークンが期限切れの場合は、再認証が必要です。

```bash
# ログを確認
tail -f logs/mail_automation.log
```

### メールが取得できない

IMAP/SMTP接続をテストしてください：

```bash
python test_imap_connection.py
```

### IMAP/SMTPが無効化されている場合

以下の代替手段を検討してください：

1. **Outlook COM（Windows限定、推奨）**
   - 基本認証不要
   - IT管理者の承認不要
   - `MAIL_PROCESSOR_TYPE=outlook_com`

2. **Microsoft Graph API**
   - OAuth2認証（より安全）
   - IT管理者によるアプリ登録が必要
   - `MAIL_PROCESSOR_TYPE=graph`

3. **Exchange Web Services (EWS)**
   - MacOS、Linux、Windows対応
   - `MAIL_PROCESSOR_TYPE=ews`

## 設定例

### 推奨設定（Windows + Outlook COM）

```bash
# メール接続
EMAIL_ADDRESS=your-email@company.com
MAIL_PROCESSOR_TYPE=outlook_com

# AI
OPENAI_API_KEY=sk-...

# 処理モード
MARK_AS_READ_AFTER_PROCESS=0

# 個人情報
MY_NAME=山田
MY_NAME_INTERNAL=山田（太郎）
MY_COMPANY_NAME=株式会社サンプル

# ドメイン
INTERNAL_DOMAINS=@company.com
GROUP_EMAIL_ADDRESSES=team@company.com

# 除外設定
EXCLUDE_SENDER_DOMAINS=@teams.mail.microsoft
EXCLUDE_AND_READ_DOMAINS=@newsletter.com

# Obsidian
OBSIDIAN_VAULT_PATH=G:\path\to\vault
```

## ライセンス

MIT License

## 作者

このプロジェクトは個人の業務効率化のために開発されました。

## 免責事項

- このツールは個人/社内使用を想定しています
- メール内容の機密性に注意してください
- OpenAI APIの利用料金が発生します
- 会社のセキュリティポリシーを確認してください

