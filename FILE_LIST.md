# 公開ファイル一覧

このディレクトリには、note.comで公開するためのファイルが含まれています。

## 📁 ディレクトリ構造

```
note/
├── README.md                     # メインドキュメント
├── INTRODUCTION.md               # システム紹介（note記事用）
├── SETUP_GUIDE.md                # セットアップガイド
├── NOTE_ARTICLE.md               # note.com記事本文
├── FILE_LIST.md                  # このファイル
├── .gitignore                    # Git除外設定
├── env.example                   # 環境変数テンプレート
├── requirements.txt              # Pythonパッケージ
│
├── main.py                       # メインプログラム
├── config.py                     # 設定管理
├── ai_processor.py               # AI処理
├── mail_processor.py             # メール処理（IMAP/SMTP）
├── outlook_com_processor.py      # メール処理（Outlook COM/Windows）
├── ews_mail_processor.py         # メール処理（EWS）
├── graph_mail_processor.py       # メール処理（Graph API）
├── storage_manager.py            # Obsidian保存
├── utils.py                      # ユーティリティ関数
│
├── setup.sh                      # セットアップスクリプト（macOS）
├── run.sh                        # 実行スクリプト（macOS）
├── install_launchagent.sh        # 自動実行設定（macOS）
│
├── templates/                    # メールテンプレート
│   ├── internal_template.txt    # 社内用返信テンプレート
│   └── external_template.txt    # 社外用返信テンプレート
│
├── logs/                         # ログディレクトリ（空）
└── credentials/                  # 認証情報ディレクトリ（空）
```

## 📄 ファイル説明

### ドキュメント

| ファイル | 説明 | 対象読者 |
|---------|------|---------|
| `README.md` | システム全体の説明 | すべてのユーザー |
| `INTRODUCTION.md` | システム紹介（詳細） | 初めてのユーザー |
| `SETUP_GUIDE.md` | 詳細なセットアップ手順 | 導入時 |
| `NOTE_ARTICLE.md` | note.com記事本文 | note読者 |

### Python プログラム

| ファイル | 行数 | 説明 |
|---------|------|------|
| `main.py` | 470+ | メイン処理、フローA/B/C |
| `config.py` | 130+ | 環境変数読み込み、設定管理 |
| `ai_processor.py` | 230+ | OpenAI API呼び出し、要約・返信生成 |
| `mail_processor.py` | 260+ | IMAP/SMTP接続 |
| `outlook_com_processor.py` | 820+ | Outlook COM接続（Windows） |
| `ews_mail_processor.py` | 250+ | Exchange Web Services接続 |
| `graph_mail_processor.py` | 260+ | Microsoft Graph API接続 |
| `storage_manager.py` | 200+ | Obsidian保存、Slack通知 |
| `utils.py` | 200+ | ログ、日時処理、フォーマット |

### 設定・スクリプト

| ファイル | 説明 |
|---------|------|
| `env.example` | 環境変数のテンプレート |
| `requirements.txt` | Pythonパッケージ一覧 |
| `.gitignore` | Git除外設定 |
| `setup.sh` | セットアップ自動化（macOS） |
| `run.sh` | 実行スクリプト（macOS） |
| `install_launchagent.sh` | 定期実行設定（macOS） |

### テンプレート

| ファイル | 説明 |
|---------|------|
| `templates/internal_template.txt` | 社内用返信テンプレート |
| `templates/external_template.txt` | 社外用返信テンプレート |

## 🚀 クイックスタート

1. すべてのファイルをダウンロード
2. `env.example`を`.env`にコピーして編集
3. `python -m venv venv` で仮想環境を作成
4. `pip install -r requirements.txt` でパッケージをインストール
5. `python main.py` で実行

詳細は `SETUP_GUIDE.md` を参照してください。

## 📊 コード統計

- **総行数**: 約2,500行
- **Python ファイル**: 9個
- **設定ファイル**: 4個
- **ドキュメント**: 4個
- **テンプレート**: 2個

## 🎯 主な機能

### メール処理
- [x] IMAP/SMTP接続
- [x] Outlook COM接続（Windows）
- [x] Graph API接続
- [x] EWS接続
- [x] 未読メール取得
- [x] 既読/未読制御
- [x] フォルダ横断検索

### AI処理
- [x] メール要約（GPT-4）
- [x] 返信文生成
- [x] 文体学習
- [x] TODO抽出
- [x] 異常検知

### 判定ロジック
- [x] To/CC/BCC対応
- [x] グループメール対応
- [x] 名前検出
- [x] 除外リスト（2種類）
- [x] 社内外判定

### テンプレート
- [x] 社内用テンプレート
- [x] 社外用テンプレート
- [x] 変数置換機能
- [x] ファイルベース管理

### Obsidian連携
- [x] Markdown保存
- [x] Frontmatter（メタデータ）
- [x] タグ自動生成
- [x] リンク自動生成
- [x] 日次インデックス
- [x] 送信者インデックス

### その他
- [x] Slack通知
- [x] 日次レポート（フローC）
- [x] ログ管理（日付別）
- [x] エラーハンドリング

## 💡 カスタマイズのヒント

### 除外設定

```bash
# 後で確認したいメール（未読維持）
EXCLUDE_SENDER_DOMAINS=@teams.mail.microsoft

# 不要なメール（既読）
EXCLUDE_AND_READ_DOMAINS=newsletter@example.com
```

### テンプレート

ファイルを編集するだけで即反映：

```
templates/internal_template.txt
templates/external_template.txt
```

### プロジェクト管理

```bash
PROJECT_KEYWORDS=プロジェクトA,プロジェクトB
```

## 📝 note.com記事での使用方法

1. `NOTE_ARTICLE.md` を記事本文として使用
2. 必要に応じて`INTRODUCTION.md`の内容を追加
3. GitHubリポジトリのURLを追記
4. スクリーンショットを追加すると◎

## 🔒 個人情報の取り扱い

このディレクトリには**個人情報は含まれていません**：

- ✅ メールアドレス: サンプル値
- ✅ API キー: サンプル値
- ✅ 名前: サンプル値（山田太郎等）
- ✅ 会社名: サンプル値
- ✅ ドメイン: サンプル値

安心して公開できます。

## 📦 配布方法

### GitHub

1. GitHubリポジトリを作成
2. このディレクトリの内容をpush
3. READMEを整備

### note.com

1. `NOTE_ARTICLE.md`の内容を記事として投稿
2. GitHubリポジトリへのリンクを追加
3. スクリーンショットを追加

### その他

- Zennで技術記事として公開
- Qiitaで共有
- 個人ブログで紹介

## 🙏 謝辞

このプロジェクトは以下の技術を使用しています：

- OpenAI GPT-4
- Microsoft Outlook
- Obsidian
- Python

## 📞 サポート

質問や問題がある場合：

1. `README.md`のトラブルシューティング参照
2. ログファイルを確認（`LOG_LEVEL=DEBUG`推奨）
3. GitHubのIssueで質問

---

**note.com公開用ファイル一式 - 2025年11月作成**

