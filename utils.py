import logging
from datetime import datetime, timedelta
import pytz
from pathlib import Path

def get_logger(name):
    """ロガーを取得（日付単位のログファイルに保存）"""
    from config import Config
    
    logger = logging.getLogger(name)
    
    # 設定からログレベルを取得
    log_level_str = Config.LOG_LEVEL
    log_level = getattr(logging, log_level_str, logging.INFO)
    logger.setLevel(log_level)
    
    if not logger.handlers:
        # 日付を含むログファイル名を生成（JST基準）
        jst = pytz.timezone('Asia/Tokyo')
        now_jst = datetime.now(jst)
        date_str = now_jst.strftime('%Y-%m-%d')
        log_filename = f'logs/mail_automation_{date_str}.log'
        
        # logsディレクトリが存在しない場合は作成
        Path('logs').mkdir(exist_ok=True)
        
        handler = logging.FileHandler(log_filename, encoding='utf-8')
        handler.setLevel(log_level)  # ハンドラーにもログレベルを設定
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        
        # コンソール出力も追加
        console_handler = logging.StreamHandler()
        console_handler.setLevel(log_level)  # ハンドラーにもログレベルを設定
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
    
    return logger

def get_jst_now():
    """JSTの現在時刻を取得"""
    jst = pytz.timezone('Asia/Tokyo')
    return datetime.now(jst)

def get_time_range_for_daily_report():
    """日次レポート用の時間範囲を取得"""
    jst = pytz.timezone('Asia/Tokyo')
    now = datetime.now(jst)
    
    # 今日の8時
    today_8am = now.replace(hour=8, minute=0, second=0, microsecond=0)
    
    # 昨日の8時
    yesterday_8am = today_8am - timedelta(days=1)
    
    return yesterday_8am, today_8am

def extract_sender_name(from_address):
    """送信者名を抽出"""
    # "山田太郎 <yamada@example.com>" の形式から名前を抽出
    import re
    match = re.match(r'^(.*?)\s*<', from_address)
    if match:
        return match.group(1).strip()
    # メールアドレスのみの場合は@より前を使用
    match = re.match(r'^([^@]+)@', from_address)
    if match:
        return match.group(1).strip()
    return from_address.strip()

def extract_project_tags(subject, body, project_keywords):
    """件名と本文からプロジェクト名を抽出してタグ化"""
    tags = []
    text = f"{subject} {body}".lower()
    
    for keyword in project_keywords:
        if keyword.strip() and keyword.strip().lower() in text:
            # タグ用にクリーンアップ（スペースを削除）
            tag = keyword.strip().replace(' ', '_')
            tags.append(tag)
    
    return tags

def sanitize_filename(name):
    """ファイル名として使用できるように文字列をサニタイズ"""
    import re
    # Windowsで使えない文字を削除
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    # 連続するスペースを1つに
    name = re.sub(r'\s+', ' ', name)
    return name.strip()

def format_markdown_email_summary(subject, from_address, summary, todos=None, 
                                  folder=None, date_received=None, message=None):
    """メール要約をMarkdown形式でフォーマット（Obsidian連携強化版）"""
    from config import Config
    
    now = datetime.now()
    date_str = now.strftime('%Y-%m-%d')
    time_str = now.strftime('%H:%M:%S')
    
    # 送信者名を抽出
    sender_name = extract_sender_name(from_address)
    sender_clean = sanitize_filename(sender_name)
    
    content = ""
    
    # === Frontmatter（メタデータ）===
    if Config.OBSIDIAN_USE_FRONTMATTER:
        frontmatter = f"""---
type: email
date: {date_str}
time: {time_str}
sender: {sender_clean}
folder: {folder or '受信トレイ'}
status: unprocessed
"""
        # プロジェクトタグを抽出
        if Config.PROJECT_KEYWORDS:
            body_text = message.get('body', '') if message else ''
            project_tags = extract_project_tags(subject, body_text, Config.PROJECT_KEYWORDS)
            if project_tags:
                frontmatter += f"projects:\n"
                for tag in project_tags:
                    frontmatter += f"  - {tag}\n"
        
        frontmatter += "---\n\n"
        content += frontmatter
    
    # === タイトル ===
    content += f"# {subject}\n\n"
    
    # === メタ情報セクション ===
    content += "## メール情報\n\n"
    
    # 送信者（リンク付き）
    if Config.OBSIDIAN_USE_LINKS:
        content += f"- **送信者**: [[送信者/{sender_clean}|{sender_name}]]\n"
    else:
        content += f"- **送信者**: {sender_name}\n"
    
    # 受信日時（日次インデックスへのリンク付き）
    if Config.OBSIDIAN_USE_LINKS:
        content += f"- **受信日時**: [[日次/{date_str}|{date_str}]] {time_str}\n"
    else:
        content += f"- **受信日時**: {date_str} {time_str}\n"
    
    # フォルダ
    if folder:
        content += f"- **フォルダ**: {folder}\n"
    
    content += "\n"
    
    # === タグセクション ===
    if Config.OBSIDIAN_USE_TAGS:
        tags = ['#メール']
        
        # 送信者タグ
        sender_tag = sender_clean.replace(' ', '_').replace('.', '_')
        tags.append(f'#送信者/{sender_tag}')
        
        # フォルダタグ
        if folder:
            folder_tag = folder.replace(' ', '_').replace('/', '_')
            tags.append(f'#フォルダ/{folder_tag}')
        
        # プロジェクトタグ
        if Config.PROJECT_KEYWORDS and message:
            body_text = message.get('body', '')
            project_tags = extract_project_tags(subject, body_text, Config.PROJECT_KEYWORDS)
            for tag in project_tags:
                tags.append(f'#プロジェクト/{tag}')
        
        content += f"{' '.join(tags)}\n\n"
    
    # === 要約セクション ===
    content += f"## 要約\n\n{summary}\n\n"
    
    # === TODOセクション ===
    if todos:
        content += f"## TODO\n\n{todos}\n\n"
    
    # === 関連リンクセクション ===
    if Config.OBSIDIAN_USE_LINKS:
        content += "## 関連\n\n"
        content += f"- [[日次/{date_str}|本日のメール一覧]]\n"
        content += f"- [[送信者/{sender_clean}|{sender_name}からのメール一覧]]\n"
    
    return content

