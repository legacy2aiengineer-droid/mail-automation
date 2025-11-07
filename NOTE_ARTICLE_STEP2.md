# 【実践編】メール自動処理システムのコード全公開 - エンジニア向けセットアップガイド

前回の記事「[ChatGPTとCursorで作った「メール自動処理システム」](https://note.com/legacy2ai/n/n49f9835b565c)」の続編です。

今回は、**GitHubで全コード（2,500行以上）を公開**し、エンジニアの方がすぐに使えるようにしました。

**GitHubリポジトリ**: https://github.com/legacy2aiengineer-droid/mail-automation

⭐ **スターをいただけると励みになります！**

---

## 📌 この記事の対象者

- ✅ Pythonの基本的な知識がある方
- ✅ ターミナル/コマンドラインに慣れている方
- ✅ Gitの基本操作ができる方
- ✅ 環境変数の設定ができる方

**💡 非エンジニアの方へ:**  
次回の記事（STEP3）で、Pythonのインストールから画面付きで詳しく解説します。そちらをお待ちください！

---

## 📖 目次

1. [前回のおさらい](#前回のおさらい)
2. [システムの全機能](#システムの全機能)
3. [メール処理の5つのパターン](#メール処理の5つのパターン)
4. [社内外テンプレート機能](#社内外テンプレート機能)
5. [セットアップ方法](#セットアップ方法)
6. [カスタマイズ例](#カスタマイズ例)
7. [実際の運用データ](#実際の運用データ)
8. [まとめ](#まとめ)

---

## 前回のおさらい

前回の記事で紹介したシステムは：

- メール処理時間を**60分→20分に削減**（67%削減）
- 重要なメールを見逃しゼロに
- ChatGPTで要約、Cursorでコード生成

今回は、このシステムを**さらに進化**させました。

### 🆕 新機能

1. **5パターンのメール分類**（除外メール対応）
2. **社内外テンプレート自動切替**
3. **グループメール・BCC対応**
4. **テンプレートファイル化**（自由にカスタマイズ可能）
5. **日次レポート自動生成**（1日1回、重複防止）

---

## システムの全機能

### ✨ コア機能

| 機能 | 説明 |
|-----|------|
| **メール自動分類** | 5パターンに自動分類（後述） |
| **AI要約** | GPT-4で長文メールを要約 |
| **返信生成** | 過去のメールから文体を学習 |
| **社内外判定** | ドメインで自動判定してテンプレート切替 |
| **Obsidian連携** | メール履歴を自動保存 |
| **TODO抽出** | メール内のタスクを自動抽出 |
| **異常検知** | 定期メールの異常をSlack通知 |
| **日次レポート** | 1日1回自動生成 |

### 🔧 対応環境

| 接続方法 | Windows | macOS | 説明 |
|---------|---------|-------|------|
| Outlook COM | ✅ | ❌ | 基本認証不要（推奨） |
| Graph API | ✅ | ✅ | OAuth2（安全） |
| EWS | ✅ | ✅ | IMAP無効でも動作 |
| IMAP/SMTP | ✅ | ✅ | 一般的な方法 |

---

## メール処理の5つのパターン

前回は「自分宛」「自分宛以外」の2パターンでしたが、今回は**5パターン**に進化しました。

### 🗑️ 除外（既読） - セールスメール対応

**条件**: 指定したドメインからのメール

```bash
EXCLUDE_AND_READ_DOMAINS=@newsletter.com,sales@example.com
```

**処理**:
- ❌ AI処理なし
- ✅ 自動で既読
- 📦 Obsidian保存なし

**用途**: セールスメール、ニュースレター、広告など

---

### 🚫 除外（未読維持） - Teams通知対応

**条件**: 指定したドメインからのメール

```bash
EXCLUDE_SENDER_DOMAINS=@teams.mail.microsoft
```

**処理**:
- ❌ AI処理なし
- ✅ 未読のまま維持
- 📦 Obsidian保存なし

**用途**: Teams通知、システム通知、後で確認したいメール

---

### 📧 パターン1: 要約+返信下書き

**条件**: Toに自分のメールアドレスが含まれる

**処理**:
- ✅ AI要約
- ✅ 返信下書き自動生成
- ✅ TODO抽出
- ✅ Obsidian保存
- ✅ 未読のまま維持

**用途**: 直接返信が必要なメール

---

### 📬 パターン2: 要約のみ

**条件**: 以下のいずれか
- CCに自分のメールアドレスが含まれる
- グループメール宛で本文に名前が含まれる
- To/CCに自分のアドレスがないが、本文に名前が含まれる（BCC対応）

**処理**:
- ✅ AI要約
- ❌ 返信生成なし
- ✅ TODO抽出
- ✅ Obsidian保存
- ✅ 未読のまま維持

**用途**: 情報共有メール、CCメール、BCC

---

### ✓ パターン0: 既読のみ

**条件**: 上記以外のメール

**処理**:
- ❌ AI処理なし
- ✅ 既読にする
- 📦 Obsidian保存なし

**用途**: 関係ないメール

---

## 社内外テンプレート機能

### 🏢 自動判定の仕組み

```
送信元アドレスをチェック
  ↓
@company.com から？
  ├─ YES → 社内用テンプレート
  └─ NO  → 社外用テンプレート
```

### 📧 社内用テンプレート

`templates/internal_template.txt`:

```
{sender_name}様

お疲れさまです、{my_name}です。

{body}

よろしくお願いいたします。
```

**設定**:
```bash
MY_NAME_INTERNAL=山田（太郎）
INTERNAL_DOMAINS=@company.com
```

**生成例**:
```
山田様

お疲れさまです、山田（太郎）です。

会議の件、承知いたしました。
明日の14時から参加させていただきます。

よろしくお願いいたします。
```

### 🏢 社外用テンプレート

`templates/external_template.txt`:

```
{sender_company}　{sender_name}様

お世話になっております。
{my_company}の{my_name}です。

{body}

以上、よろしくお願いいたします。
```

**設定**:
```bash
MY_NAME=山田
MY_COMPANY_NAME=株式会社サンプル
```

**生成例**:
```
ABC株式会社　田中様

お世話になっております。
株式会社テックソリューションの山田です。

お見積もりの件、承知いたしました。
明日中に資料をお送りいたします。

以上、よろしくお願いいたします。
```

### 🎨 テンプレートは自由にカスタマイズ可能

ファイルを編集するだけで、次回実行時から反映されます。

---

## セットアップ方法（エンジニア向け）

### 📦 クイックスタート（5分）

```bash
# クローン
git clone https://github.com/legacy2aiengineer-droid/mail-automation.git
cd mail-automation

# 環境設定
cp env.example .env
vim .env  # または code .env

# 依存関係
python -m venv venv
source venv/bin/activate  # macOS/Linux
# または .\venv\Scripts\Activate.ps1  # Windows
pip install -r requirements.txt

# 実行
python main.py
```

### 📋 必要な環境変数（最小構成）

```bash
EMAIL_ADDRESS=your@company.com
MAIL_PROCESSOR_TYPE=outlook_com  # Windows推奨
OPENAI_API_KEY=sk-proj-...
OBSIDIAN_VAULT_PATH=/path/to/vault
MARK_AS_READ_AFTER_PROCESS=0
MY_NAME=山田
INTERNAL_DOMAINS=@company.com
```

### 🔧 詳細設定（env.example参照）

**推奨設定（フル機能）:**

```bash
# メール接続
EMAIL_ADDRESS=you@company.com
MAIL_PROCESSOR_TYPE=outlook_com

# AI
OPENAI_API_KEY=sk-proj-...

# ルールベース判定
MARK_AS_READ_AFTER_PROCESS=0
MY_NAME=山田
MY_NAME_INTERNAL=山田（太郎）
MY_COMPANY_NAME=株式会社サンプル

# ドメイン設定
INTERNAL_DOMAINS=@company.com
GROUP_EMAIL_ADDRESSES=team@company.com,project@company.com

# 除外設定
EXCLUDE_SENDER_DOMAINS=@teams.mail.microsoft
EXCLUDE_AND_READ_DOMAINS=@newsletter.com,sales@example.com

# Obsidian
OBSIDIAN_VAULT_PATH=/path/to/vault
OBSIDIAN_USE_TAGS=true
OBSIDIAN_USE_LINKS=true
OBSIDIAN_CREATE_DAILY_INDEX=true

# ログレベル
LOG_LEVEL=INFO  # DEBUG | INFO | WARNING | ERROR
```

### 🎨 テンプレートのカスタマイズ

`templates/`内のファイルを編集するだけ：

```python
# templates/internal_template.txt
{sender_name}様

お疲れさまです、{my_name}です。

{body}

よろしくお願いいたします。
```

変数置換はPythonの`str.format()`を使用。

### 🚀 実行とデプロイ

```bash
# 手動実行
python main.py

# 自動実行（macOS）
chmod +x install_launchagent.sh
./install_launchagent.sh

# 自動実行（Windows）
# タスクスケジューラで設定（詳細はSTEP3記事参照）
```

---

## カスタマイズ例

### 例1: セールスメールを自動除外

```bash
EXCLUDE_AND_READ_DOMAINS=@newsletter.com,@sales.example.com,@marketing.example.com
```

これで、これらのドメインからのメールは自動で既読になります。

### 例2: グループメールの名前判定

```bash
GROUP_EMAIL_ADDRESSES=team@company.com,project@company.com
MY_NAME=山田
```

グループメール宛でも、本文に「山田」があれば要約されます。

### 例3: よりカジュアルな社内用テンプレート

`templates/internal_template.txt`:

```
{sender_name}さん

お疲れ様です！{my_name}です。

{body}

よろしくです！
```

---

## 実際の運用データ

### 📊 1週間の運用結果

**処理したメール**: 250通

| パターン | 件数 | 割合 | 処理内容 |
|---------|-----|------|---------|
| 除外（既読） | 95通 | 38% | セールスメール等 |
| パターン2（要約のみ） | 80通 | 32% | CCメール |
| パターン1（要約+返信） | 45通 | 18% | 直接返信が必要 |
| 除外（未読） | 30通 | 12% | Teams通知 |

### ⏱️ 時間の内訳

**導入前（60分/日）:**
- メールチェック: 20分
- 内容確認: 25分
- 返信作成: 15分

**導入後（20分/日）:**
- 未読メール確認: 10分（要約を読むだけ）
- 返信確認・修正: 8分（AIが生成した下書きを確認）
- システム確認: 2分

**削減時間: 40分/日 = 約14時間/月**

### 💰 コスト

- OpenAI API: 約800円/月
- セットアップ時間: 30分（初回のみ）
- **ROI（投資対効果）**: 時給換算で十分に元が取れる

---

## よくある質問

### Q1: プログラミング知識は必要？

**A**: 基本的なPythonの知識があれば十分です。セットアップはコピー&ペーストで完了します。

### Q2: 会社で使っても大丈夫？

**A**: Windows + Outlook COMを使えば、既存のOutlook設定を利用するため、IT部門の承認が不要な場合が多いです。ただし、必ず会社のポリシーを確認してください。

### Q3: 誤って重要なメールを既読にしたら？

**A**: ルールベース判定（`MARK_AS_READ_AFTER_PROCESS=0`）を使えば、重要なメールは必ず未読のまま維持されます。また、すべて未読維持する設定（`=1`）も可能です。

### Q4: OpenAI APIの料金が心配...

**A**: GPT-4を使用しますが、メール1通あたり約5〜10円程度です。1日10通処理しても月1,500〜3,000円程度です。削減される時間を考えれば十分に元が取れます。

### Q5: MacでもWindows でも使える？

**A**: はい、両方対応しています。Windowsなら**Outlook COM**、macOSなら**Graph API**または**EWS**が推奨です。

---

## 技術的なポイント

### 判定ロジックの実装

メール処理の判定は、以下の優先順位で行われます：

```python
1. 除外リストチェック（最優先）
   - セールスメール → 既読
   - Teams通知 → 未読維持

2. グループメールチェック
   - 本文に名前あり → 要約のみ
   - 本文に名前なし → 既読

3. To/CCチェック
   - To: 自分宛 → 要約+返信
   - CC: 自分宛 → 要約のみ

4. BCC対応（名前チェック）
   - 本文/件名に名前 → 要約のみ

5. デフォルト
   - 既読にする
```

### テンプレートシステム

Python の文字列フォーマット機能を使用：

```python
template = "{sender_name}様\n\n{body}\n\n{my_name}"
result = template.format(
    sender_name="山田太郎",
    body=ai_generated_content,
    my_name="山田（太郎）"
)
```

ファイルベースなので、コードを触らずにカスタマイズ可能。

### フローC（日次レポート）の改善

前回は「8時台のみ実行」でしたが、今回は：

- **8時以降の初回実行時**に実行
- **1日1回のみ**実行（重複防止）
- 実行日を記録ファイルで管理

---

## セットアップ方法（詳細）

### 📋 必要なもの

- Windows 10以降 または macOS 10.14以降
- Python 3.11以降
- OpenAI APIキー
- Outlookアカウント

### 🔧 基本設定（5分）

```bash
# 1. クローン
git clone https://github.com/legacy2aiengineer-droid/mail-automation.git
cd mail-automation

# 2. 環境設定
cp env.example .env
# .envを編集

# 3. 依存パッケージ
pip install -r requirements.txt

# 4. 実行
python main.py
```

### ⚙️ 推奨設定

```bash
# メール接続（Windows推奨）
EMAIL_ADDRESS=your-email@company.com
MAIL_PROCESSOR_TYPE=outlook_com

# AI
OPENAI_API_KEY=sk-...

# ルールベース判定
MARK_AS_READ_AFTER_PROCESS=0
MY_NAME=山田
MY_NAME_INTERNAL=山田（太郎）
MY_COMPANY_NAME=株式会社サンプル

# ドメイン
INTERNAL_DOMAINS=@company.com
GROUP_EMAIL_ADDRESSES=team@company.com

# 除外設定（重要！）
EXCLUDE_SENDER_DOMAINS=@teams.mail.microsoft
EXCLUDE_AND_READ_DOMAINS=@newsletter.com

# Obsidian
OBSIDIAN_VAULT_PATH=G:\path\to\vault
```

---

## カスタマイズ例

### 🎨 テンプレートのカスタマイズ

**もっとカジュアルに:**

```
{sender_name}さん

こんにちは！{my_name}です。

{body}

ではでは！
```

**もっとフォーマルに:**

```
{sender_company}
{sender_name} 様

平素より大変お世話になっております。
{my_company}の{my_name}でございます。

{body}

何卒よろしくお願い申し上げます。
```

### 🔧 判定ルールのカスタマイズ

**シンプルに全部未読:**

```bash
MARK_AS_READ_AFTER_PROCESS=1
```

全メールを要約+返信生成（判定なし）

**細かく制御:**

```bash
MARK_AS_READ_AFTER_PROCESS=0
MY_NAME=山田
GROUP_EMAIL_ADDRESSES=team@company.com,ml@company.com
EXCLUDE_SENDER_DOMAINS=@teams.mail.microsoft,@github.com
EXCLUDE_AND_READ_DOMAINS=@sales.com,@newsletter.com
```

---

## 実際の運用データ

### 📈 処理メールの内訳（1週間）

```
総処理数: 250通

除外（既読）        ████████████████████████ 95通（38%）
パターン2（要約のみ）  ██████████████████ 80通（32%）
パターン1（要約+返信）  ████████ 45通（18%）
除外（未読）        ██████ 30通（12%）
```

### ⏱️ 削減効果

**月間削減時間: 14時間**

これは、1日1.5冊のビジネス書を読める時間に相当します。

### 💡 具体的な使用例

**ケース1: 上司からの直接メール**
```
From: boss@company.com
To: you@company.com
Subject: プロジェクト進捗確認

→ 📧 要約+返信下書き作成
→ 未読のまま
→ Obsidianに保存
```

**ケース2: チームCC**
```
From: team-member@company.com
To: team@company.com
CC: you@company.com
本文: 「山田さん、確認お願いします」

→ 📬 要約のみ
→ 未読のまま
→ Obsidianに保存
```

**ケース3: セールスメール**
```
From: sales@newsletter.com

→ 🗑️ 既読にする
→ AI処理なし
```

---

## まとめ

### ✨ このツールで実現できること

- ✅ メール処理時間を**67%削減**（60分→20分）
- ✅ 重要なメールを**見逃しゼロ**
- ✅ 社内外で**自動的にトーンを切替**
- ✅ セールスメールを**自動で既読**
- ✅ Obsidianで**メール履歴を一元管理**

### 💰 投資対効果

| 項目 | 金額/時間 |
|-----|----------|
| 初期投資 | 30分（セットアップ） |
| 月間コスト | 500〜1,000円（OpenAI API） |
| 削減時間 | 月14時間 |
| **ROI** | **時給換算で十分に元が取れる** |

### 🎯 おすすめの人

- ✅ 毎日30通以上のメールを受信する人
- ✅ CCメールの処理に迷っている人
- ✅ Obsidianでナレッジ管理をしている人
- ✅ 業務効率化に興味がある人

---

## リンク

- **GitHubリポジトリ**: https://github.com/legacy2aiengineer-droid/mail-automation
- **前回の記事（STEP1）**: https://note.com/legacy2ai/n/n49f9835b565c
- **詳細ドキュメント**: READMEとSETUP_GUIDEをご参照ください

## フィードバック募集

このツールを使ってみて、ご意見・ご要望があれば：
- GitHubのIssueで報告
- この記事のコメント欄でコメント
- GitHubでスター⭐をいただけると励みになります！

---

## トラブルシューティング（エンジニア向け）

### Issue 1: IMAP接続エラー

```python
# test_imap_connection.py で診断
python test_imap_connection.py
```

### Issue 2: Outlook COM接続エラー

```python
# pywin32のインストール確認
pip install --upgrade pywin32
python -c "import win32com.client; print('OK')"
```

### Issue 3: OpenAI API エラー

```python
# タイムアウト設定を確認（ai_processor.py）
timeout = httpx.Timeout(60.0, connect=10.0)
```

---

## コントリビューション募集

このプロジェクトは、皆さんの貢献を歓迎します：

- 🐛 バグ報告: GitHubのIssue
- 💡 機能提案: GitHubのIssue
- 🔧 プルリクエスト: 大歓迎！
- 📝 ドキュメント改善: Typoやわかりにくい箇所の修正

---

## 次回予告（STEP3）

次回は、**非エンジニア向けの完全導入ガイド**を公開します：

- 🖼️ **画像付き詳細解説**（50枚以上のスクリーンショット）
- 🐍 **Pythonのインストール**から始める完全ガイド
- ⚙️ **環境変数の設定**を画面付きで解説
- 🔧 **よくあるエラーと解決方法**
- 📞 **サポート付き**（コメント欄で質問対応）

**STEP3は有料記事として公開予定です。**  
Pythonに触ったことがない方でも、画面を見ながら設定できる内容にします。

ぜひフォローして、続編をお待ちください！

---

### タグ
#Python #メール自動化 #OpenAI #GPT4 #Outlook #Obsidian #業務効率化 #AI活用 #ChatGPT #Cursor #エンジニア

---

**この記事が役に立ったら、ぜひスキやフォローをお願いします！**

**前回の記事（STEP1）**: https://note.com/legacy2ai/n/n49f9835b565c  
**GitHubリポジトリ**: https://github.com/legacy2aiengineer-droid/mail-automation

