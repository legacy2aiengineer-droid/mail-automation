"""
Microsoft Graph APIを使用したメール処理
OAuth2認証を使用（基本認証より安全）
"""
import os
import requests
from datetime import datetime, timedelta
from msal import ConfidentialClientApplication
from config import Config
from utils import get_logger

logger = get_logger(__name__)

class GraphMailProcessor:
    """Microsoft Graph APIを使用したメール処理"""
    
    # Microsoft Graph API エンドポイント
    GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'
    
    # 必要なスコープ
    SCOPES = ['https://graph.microsoft.com/Mail.Read',
              'https://graph.microsoft.com/Mail.Send',
              'https://graph.microsoft.com/User.Read']
    
    def __init__(self):
        self.client_id = os.getenv('AZURE_CLIENT_ID')
        self.tenant_id = os.getenv('AZURE_TENANT_ID')
        self.client_secret = os.getenv('AZURE_CLIENT_SECRET')
        self.redirect_uri = os.getenv('AZURE_REDIRECT_URI', 'http://localhost:8080/callback')
        self.authority = f'https://login.microsoftonline.com/{self.tenant_id}'
        
        if not all([self.client_id, self.tenant_id, self.client_secret]):
            raise ValueError("Azure AD設定が不完全です。.envファイルを確認してください。")
        
        # MSALアプリケーションの作成
        self.app = ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=self.authority
        )
        
        self.access_token = None
        self._acquire_token()
    
    def _acquire_token(self):
        """アクセストークンを取得"""
        try:
            # クライアントクレデンシャルフローを使用（アプリケーション権限の場合）
            # または、デバイスコードフロー、認証コードフローなどを使用
            result = self.app.acquire_token_for_client(scopes=['https://graph.microsoft.com/.default'])
            
            if 'access_token' in result:
                self.access_token = result['access_token']
                logger.info("アクセストークンの取得に成功")
            else:
                error = result.get('error_description', result.get('error', 'Unknown error'))
                logger.error(f"アクセストークンの取得に失敗: {error}")
                raise Exception(f"トークン取得エラー: {error}")
                
        except Exception as e:
            logger.error(f"認証エラー: {e}")
            raise
    
    def _get_headers(self):
        """APIリクエスト用のヘッダーを取得"""
        if not self.access_token:
            self._acquire_token()
        
        return {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
    
    def get_messages(self, unread_only=True, since_time=None):
        """メールを取得"""
        messages = []
        
        try:
            # 起動時刻から24時間前をデフォルトとして設定
            if since_time is None:
                since_time = datetime.now() - timedelta(hours=24)
                logger.info(f"デフォルト時間範囲: 過去24時間")
            
            # Graph APIのフィルタを構築
            filters = []
            
            if unread_only:
                filters.append("isRead eq false")
            
            if since_time:
                # ISO 8601形式に変換
                since_str = since_time.strftime('%Y-%m-%dT%H:%M:%SZ')
                filters.append(f"receivedDateTime ge {since_str}")
                logger.info(f"メール取得開始: {since_time.strftime('%Y-%m-%d %H:%M:%S')}以降（現在: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}）")
            
            filter_query = " and ".join(filters) if filters else None
            
            # エンドポイント
            endpoint = f"{self.GRAPH_API_ENDPOINT}/me/messages"
            params = {
                '$filter': filter_query,
                '$orderby': 'receivedDateTime desc',
                '$top': 50  # 一度に取得する最大件数
            }
            
            if not filter_query:
                params.pop('$filter')
            
            response = requests.get(endpoint, headers=self._get_headers(), params=params)
            response.raise_for_status()
            
            data = response.json()
            
            for msg in data.get('value', []):
                # To フィールドを取得
                to_recipients = msg.get('toRecipients', [])
                to_addresses = [recip.get('emailAddress', {}).get('address', '') for recip in to_recipients]
                to_field = "; ".join(to_addresses)
                
                messages.append({
                    'id': msg.get('id'),
                    'subject': msg.get('subject', ''),
                    'from': msg.get('from', {}).get('emailAddress', {}).get('address', ''),
                    'to': to_field,  # To フィールドを追加
                    'body': msg.get('bodyPreview', '') or msg.get('body', {}).get('content', ''),
                    'date': msg.get('receivedDateTime', ''),
                    'unread': not msg.get('isRead', False),
                    'folder': 'Inbox'  # Graph APIでは受信トレイから取得
                })
            
            logger.info(f"{len(messages)}件のメールを取得しました")
            
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 401:
                logger.error("認証エラー: アクセストークンを更新してください")
                self._acquire_token()
                raise
            else:
                logger.error(f"メール取得エラー: {e.response.status_code} - {e.response.text}")
                raise
        except Exception as e:
            logger.error(f"メール取得エラー: {e}")
            raise
        
        return messages
    
    def is_target_mail(self, message):
        """ターゲットメールかどうかを判定"""
        body = message.get('body', '').lower()
        return any(keyword.lower() in body for keyword in Config.KEYWORDS)
    
    def mark_as_read(self, message_id):
        """メールを既読にする"""
        try:
            endpoint = f"{self.GRAPH_API_ENDPOINT}/me/messages/{message_id}"
            data = {
                'isRead': True
            }
            
            response = requests.patch(endpoint, headers=self._get_headers(), json=data)
            response.raise_for_status()
            
            logger.info(f"メールを既読にしました: {message_id}")
        except Exception as e:
            logger.error(f"既読処理エラー: {e}")
            raise
    
    def mark_as_unread(self, message_id):
        """メールを未読にする"""
        try:
            endpoint = f"{self.GRAPH_API_ENDPOINT}/me/messages/{message_id}"
            data = {
                'isRead': False
            }
            
            response = requests.patch(endpoint, headers=self._get_headers(), json=data)
            response.raise_for_status()
            
            logger.info(f"メールを未読にしました: {message_id}")
        except Exception as e:
            logger.error(f"未読処理エラー: {e}")
            raise
    
    def send_email(self, subject, body, to_address):
        """メールを送信"""
        try:
            endpoint = f"{self.GRAPH_API_ENDPOINT}/me/sendMail"
            
            message = {
                'message': {
                    'subject': subject,
                    'body': {
                        'contentType': 'Text',
                        'content': body
                    },
                    'toRecipients': [
                        {
                            'emailAddress': {
                                'address': to_address
                            }
                        }
                    ]
                }
            }
            
            response = requests.post(endpoint, headers=self._get_headers(), json=message)
            response.raise_for_status()
            
            logger.info(f"メール送信成功: {subject}")
        except Exception as e:
            logger.error(f"メール送信エラー: {e}")
            raise
    
    def get_sent_messages(self, limit=5):
        """送信済みメールを取得（文体学習用）"""
        messages = []
        
        try:
            headers = {
                'Authorization': f'Bearer {self.access_token}',
                'Content-Type': 'application/json'
            }
            
            # 送信済みアイテムフォルダからメールを取得
            url = f"{self.GRAPH_API_ENDPOINT}/me/mailFolders/SentItems/messages"
            params = {
                '$top': limit * 2,  # 返信メールを除外するため多めに取得
                '$orderby': 'sentDateTime desc',
                '$select': 'subject,body,sentDateTime'
            }
            
            response = requests.get(url, headers=headers, params=params)
            response.raise_for_status()
            
            items = response.json().get('value', [])
            
            for item in items:
                if len(messages) >= limit:
                    break
                
                subject = item.get('subject', '')
                
                # 返信・転送メールは除外（オリジナルのメールのみ）
                if subject.startswith('Re:') or subject.startswith('FW:') or subject.startswith('Fw:'):
                    continue
                
                body_content = item.get('body', {})
                body = body_content.get('content', '') if isinstance(body_content, dict) else ''
                
                messages.append({
                    'subject': subject,
                    'body': body,
                    'date': item.get('sentDateTime', '')
                })
            
            logger.debug(f"送信済みメール {len(messages)}件を取得しました")
            return messages
            
        except Exception as e:
            logger.error(f"送信済みメール取得エラー: {e}")
            return []
    
    def disconnect(self):
        """接続を閉じる（Graph APIの場合は特に処理不要）"""
        self.access_token = None
        logger.info("Graph API接続を閉じました")

