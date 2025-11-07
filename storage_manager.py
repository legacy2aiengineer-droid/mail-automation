import os
import re
from datetime import datetime
from pathlib import Path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from config import Config
from utils import get_logger

logger = get_logger(__name__)

def sanitize_filename(filename):
    """ファイル名から不正な文字を削除"""
    # Windowsで使用できない文字を置換
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(invalid_chars, '_', filename)
    # 連続するアンダースコアを1つに
    sanitized = re.sub(r'_+', '_', sanitized)
    # 前後の空白やアンダースコアを削除
    sanitized = sanitized.strip(' _')
    # ファイル名が長すぎる場合は切り詰め（拡張子を除いて200文字まで）
    if len(sanitized) > 200:
        sanitized = sanitized[:200]
    return sanitized

class StorageManager:
    def __init__(self):
        self.drive_service = None
        self._init_drive_service()
    
    def _init_drive_service(self):
        """Google Driveサービスの初期化"""
        SCOPES = ['https://www.googleapis.com/auth/drive.file']
        creds = None
        
        if os.path.exists(Config.GOOGLE_CREDENTIALS_FILE):
            flow = InstalledAppFlow.from_client_secrets_file(
                Config.GOOGLE_CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
            self.drive_service = build('drive', 'v3', credentials=creds)
    
    def save_to_obsidian(self, content, filename, folder='Mail', sender_name=None):
        """Obsidianに保存"""
        import re
        from utils import extract_sender_name, sanitize_filename as sanitize_util
        
        # ファイル名から不正な文字を削除（Windows/macOS対応）
        # 不正な文字: \ / : * ? " < > |
        sanitized_filename = re.sub(r'[\\/:*?"<>|]', '_', filename)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        folder_path = Config.OBSIDIAN_VAULT_PATH / '01_Index' / folder
        
        # フォルダが存在しない場合は作成
        folder_path.mkdir(parents=True, exist_ok=True)
        
        file_path = folder_path / f"{timestamp}_{sanitized_filename}.md"
        
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            logger.info(f"Obsidianに保存: {file_path}")
            
            # 日次インデックスを更新
            if Config.OBSIDIAN_CREATE_DAILY_INDEX:
                self._update_daily_index(timestamp, sanitized_filename, sender_name)
            
            # 送信者インデックスを更新
            if Config.OBSIDIAN_CREATE_SENDER_INDEX and sender_name:
                self._update_sender_index(sender_name, timestamp, sanitized_filename)
            
            return file_path
        except Exception as e:
            logger.error(f"Obsidian保存エラー: {e}")
            raise
    
    def _update_daily_index(self, timestamp, filename, sender_name=None):
        """日次インデックスを更新"""
        try:
            date_str = timestamp[:8]  # YYYYMMDDの部分
            formatted_date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
            
            index_folder = Config.OBSIDIAN_VAULT_PATH / '01_Index' / 'Mail' / '日次'
            index_folder.mkdir(parents=True, exist_ok=True)
            
            index_file = index_folder / f"{formatted_date}.md"
            
            # 既存のインデックスを読み込むか、新規作成
            if index_file.exists():
                with open(index_file, 'r', encoding='utf-8') as f:
                    content = f.read()
            else:
                content = f"""# {formatted_date} メール一覧

#日次メール #メール

## 受信メール

"""
            
            # 新しいメールのリンクを追加
            time_part = timestamp[9:15]  # HHMMSS
            formatted_time = f"{time_part[:2]}:{time_part[2:4]}"
            sender_info = f" - {sender_name}" if sender_name else ""
            
            link_line = f"- {formatted_time}{sender_info} [[../Mail/{timestamp}_{filename}|{filename}]]\n"
            
            # 重複チェック
            if link_line not in content:
                content += link_line
            
            # インデックスを保存
            with open(index_file, 'w', encoding='utf-8') as f:
                f.write(content)
            
            logger.debug(f"日次インデックス更新: {index_file}")
        except Exception as e:
            logger.warning(f"日次インデックス更新エラー: {e}")
    
    def _update_sender_index(self, sender_name, timestamp, filename):
        """送信者別インデックスを更新"""
        try:
            from utils import sanitize_filename as sanitize_util
            
            sender_clean = sanitize_util(sender_name)
            
            index_folder = Config.OBSIDIAN_VAULT_PATH / '01_Index' / 'Mail' / '送信者'
            index_folder.mkdir(parents=True, exist_ok=True)
            
            index_file = index_folder / f"{sender_clean}.md"
            
            # 既存のインデックスを読み込むか、新規作成
            if index_file.exists():
                with open(index_file, 'r', encoding='utf-8') as f:
                    content = f.read()
            else:
                content = f"""# {sender_name} からのメール

#送信者/{sender_clean.replace(' ', '_')} #メール

## メール一覧

"""
            
            # 新しいメールのリンクを追加
            date_part = timestamp[:8]  # YYYYMMDD
            time_part = timestamp[9:15]  # HHMMSS
            formatted_datetime = f"{date_part[:4]}-{date_part[4:6]}-{date_part[6:8]} {time_part[:2]}:{time_part[2:4]}"
            
            link_line = f"- {formatted_datetime} [[../Mail/{timestamp}_{filename}|{filename}]]\n"
            
            # 重複チェック
            if link_line not in content:
                content += link_line
            
            # インデックスを保存
            with open(index_file, 'w', encoding='utf-8') as f:
                f.write(content)
            
            logger.debug(f"送信者インデックス更新: {index_file}")
        except Exception as e:
            logger.warning(f"送信者インデックス更新エラー: {e}")
    
    def save_attachments_to_drive(self, attachments, message_id):
        """添付ファイルをGoogle Driveに保存"""
        if not self.drive_service:
            logger.warning("Google Driveサービスが初期化されていません")
            return []
        
        saved_files = []
        for attachment in attachments:
            # 添付ファイルをダウンロード
            # (実装は添付ファイルの取得方法に依存)
            saved_files.append(attachment)
        
        return saved_files
    
    def send_slack_notification(self, message):
        """Slackに通知"""
        import requests
        
        url = "https://slack.com/api/chat.postMessage"
        headers = {
            'Authorization': f'Bearer {Config.SLACK_BOT_TOKEN}',
            'Content-Type': 'application/json'
        }
        payload = {
            'channel': Config.SLACK_CHANNEL,
            'text': message
        }
        
        response = requests.post(url, headers=headers, json=payload)
        logger.info(f"Slack通知: {response.status_code}")

