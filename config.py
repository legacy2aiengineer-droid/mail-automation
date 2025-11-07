import os
from pathlib import Path
from dotenv import load_dotenv

# .envファイルの読み込み
load_dotenv()

class Config:
    # IMAP/SMTP設定
    EMAIL_ADDRESS = os.getenv('EMAIL_ADDRESS')
    EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
    IMAP_SERVER = os.getenv('IMAP_SERVER', 'outlook.office365.com')
    IMAP_PORT = int(os.getenv('IMAP_PORT', '993'))
    SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.office365.com')
    SMTP_PORT = int(os.getenv('SMTP_PORT', '587'))
    
    # OpenAI API
    OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')
    
    # Slack
    SLACK_BOT_TOKEN = os.getenv('SLACK_BOT_TOKEN')
    SLACK_CHANNEL = os.getenv('SLACK_CHANNEL', '#monitoring')
    
    # Google Drive
    GOOGLE_CREDENTIALS_FILE = os.getenv('GOOGLE_CREDENTIALS_FILE')
    GOOGLE_DRIVE_FOLDER_ID = os.getenv('GOOGLE_DRIVE_FOLDER_ID')
    
    # Obsidian (Windows用のパス設定)
    obsidian_path = os.getenv('OBSIDIAN_VAULT_PATH')
    if obsidian_path:
        OBSIDIAN_VAULT_PATH = Path(obsidian_path)
        OBSIDIAN_MAIL_FOLDER = OBSIDIAN_VAULT_PATH / '01_Index' / 'Mail'
        OBSIDIAN_TODO_FOLDER = OBSIDIAN_VAULT_PATH / '04_Templates'
    else:
        # デフォルトパス（環境変数が設定されていない場合）
        OBSIDIAN_VAULT_PATH = Path('.')
        OBSIDIAN_MAIL_FOLDER = OBSIDIAN_VAULT_PATH / 'obsidian' / 'Mail'
        OBSIDIAN_TODO_FOLDER = OBSIDIAN_VAULT_PATH / 'obsidian' / 'TODO'
    
    # ターゲットメール
    TARGET_EMAIL = os.getenv('TARGET_EMAIL')
    
    # キーワード
    KEYWORDS = ['カナデビア', '中村']
    NORMAL_KEYWORDS = ['OK', 'No Error', 'Healthy', '200']
    ABNORMAL_KEYWORDS = ['ERROR', 'FAIL', 'CRITICAL', 'Timeout']
    
    # SMTP送信設定（SMTP認証が無効な場合はfalseに設定）
    ENABLE_EMAIL_SENDING = os.getenv('ENABLE_EMAIL_SENDING', 'true').lower() == 'true'
    
    # メールを下書きに保存するか送信するか（true=下書き保存、false=送信）
    SAVE_AS_DRAFT = os.getenv('SAVE_AS_DRAFT', 'true').lower() == 'true'
    
    # メール処理後の既読設定（0=ルールベース、1=全部未読、2=全部既読）
    # 0: Toに自分のアドレス/名前が含まれる → 未読、それ以外（CC等） → 既読
    # 1: すべて未読のまま維持
    # 2: すべて既読にする
    _mark_as_read_env = os.getenv('MARK_AS_READ_AFTER_PROCESS', '1')
    try:
        MARK_AS_READ_AFTER_PROCESS = int(_mark_as_read_env)
        if MARK_AS_READ_AFTER_PROCESS not in [0, 1, 2]:
            print(f"警告: MARK_AS_READ_AFTER_PROCESS の値が範囲外です({_mark_as_read_env})。デフォルト値(1)を使用します。")
            MARK_AS_READ_AFTER_PROCESS = 1
    except (ValueError, TypeError):
        print(f"警告: MARK_AS_READ_AFTER_PROCESS の値が不正です({_mark_as_read_env})。デフォルト値(1)を使用します。")
        MARK_AS_READ_AFTER_PROCESS = 1
    
    # 自分の名前（ルールベース判定用）
    MY_NAME = os.getenv('MY_NAME', '')
    
    # 返信テンプレート用の設定
    MY_NAME_INTERNAL = os.getenv('MY_NAME_INTERNAL', '')  # 社内用の名前（例: 中村（康））
    MY_COMPANY_NAME = os.getenv('MY_COMPANY_NAME', '')  # 会社名（例: 株式会社カナデビア）
    INTERNAL_DOMAINS = os.getenv('INTERNAL_DOMAINS', '').split(',') if os.getenv('INTERNAL_DOMAINS') else []  # 社内ドメイン
    
    # テンプレートファイルのパス
    INTERNAL_EMAIL_TEMPLATE = os.getenv('INTERNAL_EMAIL_TEMPLATE', 'templates/internal_template.txt')
    EXTERNAL_EMAIL_TEMPLATE = os.getenv('EXTERNAL_EMAIL_TEMPLATE', 'templates/external_template.txt')
    
    # グループメールアドレス（カンマ区切り）
    # Toにこれらのアドレスが含まれる場合も「自分宛」として扱う
    GROUP_EMAIL_ADDRESSES = os.getenv('GROUP_EMAIL_ADDRESSES', '').split(',') if os.getenv('GROUP_EMAIL_ADDRESSES') else []
    
    # 除外する送信元ドメイン（カンマ区切り）
    # これらのドメインからのメールはAI処理をスキップして未読のまま維持
    EXCLUDE_SENDER_DOMAINS = os.getenv('EXCLUDE_SENDER_DOMAINS', '').split(',') if os.getenv('EXCLUDE_SENDER_DOMAINS') else []
    
    # 除外して既読にする送信元ドメイン（カンマ区切り）
    # これらのドメインからのメールはAI処理をスキップして既読にする（セールスメールなど）
    EXCLUDE_AND_READ_DOMAINS = os.getenv('EXCLUDE_AND_READ_DOMAINS', '').split(',') if os.getenv('EXCLUDE_AND_READ_DOMAINS') else []
    
    # Obsidian連携設定
    OBSIDIAN_USE_TAGS = os.getenv('OBSIDIAN_USE_TAGS', 'true').lower() == 'true'
    OBSIDIAN_USE_LINKS = os.getenv('OBSIDIAN_USE_LINKS', 'true').lower() == 'true'
    OBSIDIAN_USE_FRONTMATTER = os.getenv('OBSIDIAN_USE_FRONTMATTER', 'true').lower() == 'true'
    OBSIDIAN_CREATE_DAILY_INDEX = os.getenv('OBSIDIAN_CREATE_DAILY_INDEX', 'true').lower() == 'true'
    OBSIDIAN_CREATE_SENDER_INDEX = os.getenv('OBSIDIAN_CREATE_SENDER_INDEX', 'false').lower() == 'true'
    
    # プロジェクト名の抽出キーワード（タグ生成用）
    PROJECT_KEYWORDS = os.getenv('PROJECT_KEYWORDS', '').split(',') if os.getenv('PROJECT_KEYWORDS') else []
    
    # Outlookフォルダ除外設定
    # 除外するフォルダ名（部分一致、カンマ区切り）
    OUTLOOK_EXCLUDE_FOLDERS = os.getenv('OUTLOOK_EXCLUDE_FOLDERS', '').split(',') if os.getenv('OUTLOOK_EXCLUDE_FOLDERS') else []
    # カレンダーフォルダを除外するか
    OUTLOOK_EXCLUDE_CALENDAR = os.getenv('OUTLOOK_EXCLUDE_CALENDAR', 'true').lower() == 'true'
    # 連絡先フォルダを除外するか
    OUTLOOK_EXCLUDE_CONTACTS = os.getenv('OUTLOOK_EXCLUDE_CONTACTS', 'true').lower() == 'true'
    # タスクフォルダを除外するか
    OUTLOOK_EXCLUDE_TASKS = os.getenv('OUTLOOK_EXCLUDE_TASKS', 'true').lower() == 'true'
    # メモフォルダを除外するか
    OUTLOOK_EXCLUDE_NOTES = os.getenv('OUTLOOK_EXCLUDE_NOTES', 'true').lower() == 'true'
    # ジャーナルフォルダを除外するか
    OUTLOOK_EXCLUDE_JOURNAL = os.getenv('OUTLOOK_EXCLUDE_JOURNAL', 'true').lower() == 'true'
    
    # Azure AD設定（Microsoft Graph API使用時）
    AZURE_CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
    AZURE_TENANT_ID = os.getenv('AZURE_TENANT_ID')
    AZURE_CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
    AZURE_REDIRECT_URI = os.getenv('AZURE_REDIRECT_URI', 'http://localhost:8080/callback')
    
    # EWS設定（Exchange Web Services使用時）
    EWS_SERVER = os.getenv('EWS_SERVER')  # 例: 'outlook.office365.com' または None（自動検出）
    EWS_USE_NTLM = os.getenv('EWS_USE_NTLM', 'false').lower() == 'true'
    
    # メールプロセッサーの選択（'imap', 'graph', 'outlook_com', 'ews'）
    MAIL_PROCESSOR_TYPE = os.getenv('MAIL_PROCESSOR_TYPE', 'imap')
    
    # 文体学習機能（送信済みメールから自分の文体を学習して返信文に反映）
    LEARN_WRITING_STYLE = os.getenv('LEARN_WRITING_STYLE', 'true').lower() == 'true'
    WRITING_STYLE_SAMPLE_COUNT = int(os.getenv('WRITING_STYLE_SAMPLE_COUNT', '5'))
    
    # ログ設定
    LOG_DIR = Path('logs')
    LOG_DIR.mkdir(exist_ok=True)
    
    # ログレベル設定（DEBUG, INFO, WARNING, ERROR）
    LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO').upper()
    # 有効なログレベルかチェック
    if LOG_LEVEL not in ['DEBUG', 'INFO', 'WARNING', 'ERROR']:
        LOG_LEVEL = 'INFO'

# Obsidianディレクトリの作成
try:
    Config.OBSIDIAN_MAIL_FOLDER.mkdir(parents=True, exist_ok=True)
    Config.OBSIDIAN_TODO_FOLDER.mkdir(parents=True, exist_ok=True)
except Exception as e:
    print(f"警告: Obsidianディレクトリの作成に失敗しました: {e}")

