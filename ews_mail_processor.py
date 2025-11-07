from exchangelib import Credentials, Account, DELEGATE, Configuration, NTLM
from exchangelib import Message, Mailbox, HTMLBody
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
from datetime import datetime, timedelta
import urllib3
from config import Config
from utils import get_logger

# SSL警告を抑制（必要に応じて）
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logger = get_logger(__name__)

class EWSMailProcessor:
    """Exchange Web Services を使用したメールプロセッサー"""
    
    def __init__(self):
        self.email = Config.EMAIL_ADDRESS
        self.password = Config.EMAIL_PASSWORD
        self.ews_server = Config.EWS_SERVER
        self.account = None
        
    def connect(self):
        """EWSアカウントに接続"""
        try:
            # 認証情報の設定
            credentials = Credentials(username=self.email, password=self.password)
            
            # サーバー設定
            if self.ews_server:
                # 手動でサーバーを指定する場合
                config = Configuration(
                    server=self.ews_server,
                    credentials=credentials,
                    auth_type=NTLM if Config.EWS_USE_NTLM else None
                )
                self.account = Account(
                    primary_smtp_address=self.email,
                    config=config,
                    autodiscover=False,
                    access_type=DELEGATE
                )
            else:
                # 自動検出を使用
                self.account = Account(
                    primary_smtp_address=self.email,
                    credentials=credentials,
                    autodiscover=True,
                    access_type=DELEGATE
                )
            
            # 接続確認
            _ = self.account.inbox.total_count
            logger.info("EWS接続成功")
            
        except Exception as e:
            logger.error(f"EWS接続エラー: {e}")
            logger.error(f"エラータイプ: {type(e).__name__}")
            logger.error("可能な原因: 認証情報の誤り、EWSが無効、サーバー設定の誤り")
            raise
    
    def get_sent_messages(self, limit=5):
        """送信済みメールを取得（文体学習用）"""
        messages = []
        
        try:
            from exchangelib import Message
            
            # 送信済みアイテムフォルダを取得
            sent_folder = self.account.sent
            
            # 最新のメールを取得（limit * 2件、返信除外のため多めに）
            items = sent_folder.all().order_by('-datetime_sent')[:limit * 2]
            
            for item in items:
                if len(messages) >= limit:
                    break
                
                try:
                    if not isinstance(item, Message):
                        continue
                    
                    subject = item.subject or ""
                    
                    # 返信・転送メールは除外
                    if subject.startswith('Re:') or subject.startswith('FW:') or subject.startswith('Fw:'):
                        continue
                    
                    body = item.text_body or item.body or ""
                    date = str(item.datetime_sent) if item.datetime_sent else ""
                    
                    messages.append({
                        'subject': subject,
                        'body': body,
                        'date': date
                    })
                    
                except Exception as e:
                    logger.debug(f"送信済みメール取得エラー: {e}")
                    continue
            
            logger.debug(f"送信済みメール {len(messages)}件を取得しました")
            return messages
            
        except Exception as e:
            logger.error(f"送信済みメール取得エラー: {e}")
            return []
    
    def disconnect(self):
        """EWSアカウントから切断"""
        try:
            if self.account:
                self.account.protocol.close()
                self.account = None
                logger.info("EWS切断")
        except Exception as e:
            logger.warning(f"切断エラー: {e}")
    
    def get_messages(self, unread_only=True, since_time=None):
        """メールを取得"""
        if not self.account:
            self.connect()
        
        messages = []
        try:
            # 起動時刻から24時間前をデフォルトとして設定
            if since_time is None:
                since_time = datetime.now() - timedelta(hours=24)
                logger.info(f"デフォルト時間範囲: 過去24時間")
            
            logger.info(f"メール取得開始: {since_time.strftime('%Y-%m-%d %H:%M:%S')}以降（現在: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}）")
            
            # 受信トレイからメールを取得
            inbox = self.account.inbox
            
            # フィルター条件を構築
            items = inbox.all()
            
            # 未読のみフィルター
            if unread_only:
                items = items.filter(is_read=False)
            
            # 日付フィルター
            if since_time:
                items = items.filter(datetime_received__gte=since_time)
            
            # 新しい順にソート
            items = items.order_by('-datetime_received')
            
            # メッセージを処理
            for item in items:
                try:
                    # 本文の取得（HTMLの場合はテキストに変換）
                    body = ""
                    if hasattr(item, 'text_body') and item.text_body:
                        body = item.text_body
                    elif hasattr(item, 'body') and item.body:
                        body = str(item.body)
                    
                    # 送信者のメールアドレス
                    from_address = ""
                    if hasattr(item, 'sender') and item.sender:
                        from_address = item.sender.email_address if hasattr(item.sender, 'email_address') else str(item.sender)
                    
                    # To フィールドを取得
                    to_field = ""
                    try:
                        if hasattr(item, 'to_recipients') and item.to_recipients:
                            to_addresses = [recip.email_address for recip in item.to_recipients if hasattr(recip, 'email_address')]
                            to_field = "; ".join(to_addresses)
                    except:
                        pass
                    
                    messages.append({
                        'id': item.id,
                        'subject': item.subject or '',
                        'from': from_address,
                        'to': to_field,  # To フィールドを追加
                        'body': body,
                        'date': item.datetime_received.isoformat() if item.datetime_received else '',
                        'folder': 'Inbox'  # EWSでは受信トレイから取得
                    })
                    
                except Exception as e:
                    logger.warning(f"メッセージ処理エラー: {e}")
                    continue
            
            logger.info(f"EWSでメールを取得: {len(messages)}件")
            
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
            if not self.account:
                self.connect()
            
            # メッセージIDからアイテムを取得
            # EWSではitem_idを直接使用
            inbox = self.account.inbox
            for item in inbox.filter(id=message_id):
                item.is_read = True
                item.save()
                logger.info(f"メールを既読にしました")
                break
                
        except Exception as e:
            logger.error(f"既読処理エラー: {e}")
    
    def mark_as_unread(self, message_id):
        """メールを未読にする"""
        try:
            if not self.account:
                self.connect()
            
            # メッセージIDからアイテムを取得
            inbox = self.account.inbox
            for item in inbox.filter(id=message_id):
                item.is_read = False
                item.save()
                logger.info(f"メールを未読にしました")
                break
                
        except Exception as e:
            logger.error(f"未読処理エラー: {e}")
            raise
    
    def send_email(self, subject, body, to_address):
        """メールを送信または下書きに保存"""
        try:
            if not self.account:
                self.connect()
            
            # メッセージの作成
            message = Message(
                account=self.account,
                subject=subject,
                body=body,
                to_recipients=[Mailbox(email_address=to_address)]
            )
            
            if Config.SAVE_AS_DRAFT:
                # 下書きフォルダに保存
                message.folder = self.account.drafts
                message.save()
                logger.info(f"下書きに保存しました: {subject}")
            else:
                # 送信
                message.send()
                logger.info(f"メール送信成功: {subject}")
                
        except Exception as e:
            logger.error(f"メール送信/保存エラー: {e}")
            logger.error(f"エラータイプ: {type(e).__name__}")
            raise

