import imaplib
import smtplib
import email
from datetime import datetime, timedelta
from email.header import decode_header
from email.mime.text import MIMEText
from config import Config
from utils import get_logger

logger = get_logger(__name__)

class MailProcessor:
    def __init__(self):
        self.email = Config.EMAIL_ADDRESS
        self.password = Config.EMAIL_PASSWORD
        self.imap_server = Config.IMAP_SERVER
        self.imap_port = Config.IMAP_PORT
        self.smtp_server = Config.SMTP_SERVER
        self.smtp_port = Config.SMTP_PORT
        self.mail = None
        
    def connect(self):
        """メールサーバーに接続"""
        try:
            self.mail = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            self.mail.login(self.email, self.password)
            self.mail.select('INBOX')
            logger.info("メールサーバーに接続成功")
        except imaplib.IMAP4.error as e:
            error_msg = str(e).lower()
            if 'login failed' in error_msg or 'authentication failed' in error_msg:
                logger.error(f"認証エラー: {e}")
                logger.error("可能な原因: パスワード間違い、アプリパスワードが必要、またはIMAPが無効化されています")
            elif 'access denied' in error_msg or 'forbidden' in error_msg:
                logger.error(f"アクセス拒否エラー: {e}")
                logger.error("会社のセキュリティポリシーでIMAPが無効化されている可能性があります")
            else:
                logger.error(f"IMAPエラー: {e}")
            raise
        except Exception as e:
            logger.error(f"メールサーバー接続エラー: {e}")
            logger.error(f"エラータイプ: {type(e).__name__}")
            raise
    
    def get_sent_messages(self, limit=5):
        """送信済みメールを取得（文体学習用）"""
        messages = []
        
        try:
            # 送信済みフォルダを選択
            # 一般的なフォルダ名を試す
            sent_folder_names = ['Sent', 'Sent Items', '送信済み', '送信済みアイテム', 'INBOX.Sent']
            sent_folder = None
            
            for folder_name in sent_folder_names:
                try:
                    self.mail.select(folder_name, readonly=True)
                    sent_folder = folder_name
                    logger.debug(f"送信済みフォルダを選択: {folder_name}")
                    break
                except:
                    continue
            
            if not sent_folder:
                logger.warning("送信済みフォルダが見つかりませんでした")
                return []
            
            # 最新のメールを取得（limit * 2件、返信除外のため多めに）
            _, message_numbers = self.mail.search(None, 'ALL')
            message_list = message_numbers[0].split()
            
            # 最新のメールから順に取得
            message_list = message_list[-(limit * 2):][::-1]
            
            for num in message_list:
                if len(messages) >= limit:
                    break
                
                try:
                    _, msg_data = self.mail.fetch(num, '(RFC822)')
                    email_body = msg_data[0][1]
                    email_message = email.message_from_bytes(email_body)
                    
                    subject = self._decode_header(email_message.get('Subject', ''))
                    
                    # 返信・転送メールは除外
                    if subject.startswith('Re:') or subject.startswith('FW:') or subject.startswith('Fw:'):
                        continue
                    
                    body = self._get_email_body(email_message)
                    date = email_message.get('Date', '')
                    
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
        """メールサーバーから切断"""
        try:
            if self.mail:
                self.mail.close()
                self.mail.logout()
                self.mail = None
                logger.info("メールサーバーから切断")
        except Exception as e:
            logger.error(f"切断エラー: {e}")
    
    def get_messages(self, unread_only=True, since_time=None):
        """メールを取得"""
        if not self.mail:
            self.connect()
        
        messages = []
        try:
            # 起動時刻から24時間前をデフォルトとして設定
            if since_time is None:
                since_time = datetime.now() - timedelta(hours=24)
                logger.info(f"デフォルト時間範囲: 過去24時間")
            
            logger.info(f"メール取得開始: {since_time.strftime('%Y-%m-%d %H:%M:%S')}以降（現在: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}）")
            
            # 検索条件を構築
            search_criteria = []
            
            if unread_only:
                search_criteria.append('UNSEEN')
            else:
                search_criteria.append('ALL')
            
            # since_timeが指定されている場合は日付フィルタを追加
            if since_time:
                # IMAP日付形式に変換 (DD-MMM-YYYY)
                date_str = since_time.strftime('%d-%b-%Y')
                search_criteria.append(f'SINCE {date_str}')
            
            search_query = ' '.join(search_criteria)
            _, msgnums = self.mail.search(None, search_query)
            
            if not msgnums[0]:
                return messages
            
            for msgnum in msgnums[0].split():
                try:
                    _, data = self.mail.fetch(msgnum, '(RFC822)')
                    if not data or not data[0]:
                        continue
                        
                    msg = email.message_from_bytes(data[0][1])
                    
                    # 件名のデコード
                    subject = ""
                    subject_header = decode_header(msg.get('Subject', ''))
                    if subject_header:
                        decoded_subject, encoding = subject_header[0]
                        if encoding:
                            subject = decoded_subject.decode(encoding)
                        else:
                            subject = decoded_subject if isinstance(decoded_subject, str) else decoded_subject.decode('utf-8', errors='ignore')
                    
                    # 送信者
                    from_address = msg.get('From', '')
                    
                    # To フィールドを取得
                    to_field = msg.get('To', '')
                    
                    # 日付
                    date_str = msg.get('Date', '')
                    
                    # 本文の取得
                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            if part.get_content_type() == "text/plain":
                                charset = part.get_content_charset() or 'utf-8'
                                try:
                                    body = part.get_payload(decode=True).decode(charset, errors='ignore')
                                    break
                                except Exception as e:
                                    logger.warning(f"本文デコードエラー: {e}")
                                    continue
                    else:
                        charset = msg.get_content_charset() or 'utf-8'
                        try:
                            body = msg.get_payload(decode=True).decode(charset, errors='ignore')
                        except Exception as e:
                            logger.warning(f"本文デコードエラー: {e}")
                    
                    messages.append({
                        'id': msgnum.decode(),
                        'subject': subject,
                        'from': from_address,
                        'to': to_field,  # To フィールドを追加
                        'body': body,
                        'date': date_str,
                        'folder': 'INBOX'  # IMAPでは受信トレイから取得
                    })
                except Exception as e:
                    logger.warning(f"メール {msgnum.decode()} の処理中にエラー: {e}")
                    continue
        
        except Exception as e:
            logger.error(f"メール取得エラー: {e}")
        
        return messages
    
    def is_target_mail(self, message):
        """ターゲットメールかどうかを判定"""
        body = message.get('body', '').lower()
        return any(keyword.lower() in body for keyword in Config.KEYWORDS)
    
    def mark_as_read(self, message_id):
        """メールを既読にする"""
        try:
            if isinstance(message_id, str):
                message_id = message_id.encode()
            self.mail.store(message_id, '+FLAGS', '\\Seen')
            logger.info(f"メールを既読にしました: {message_id.decode() if isinstance(message_id, bytes) else message_id}")
        except Exception as e:
            logger.error(f"既読処理エラー: {e}")
    
    def mark_as_unread(self, message_id):
        """メールを未読にする"""
        try:
            if isinstance(message_id, str):
                message_id = message_id.encode()
            self.mail.store(message_id, '-FLAGS', '\\Seen')
            logger.info(f"メールを未読にしました: {message_id.decode() if isinstance(message_id, bytes) else message_id}")
        except Exception as e:
            logger.error(f"未読処理エラー: {e}")
            raise
    
    def send_email(self, subject, body, to_address):
        """メールを送信"""
        try:
            msg = MIMEText(body, 'plain', 'utf-8')
            msg['Subject'] = subject
            msg['From'] = self.email
            msg['To'] = to_address
            
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.email, self.password)
                server.send_message(msg)
            
            logger.info(f"メール送信成功: {subject}")
        except smtplib.SMTPAuthenticationError as e:
            error_msg = str(e).lower()
            if 'smtpclientauthentication is disabled' in error_msg or 'smtp_auth_disabled' in error_msg:
                logger.error(f"SMTP認証エラー: {e}")
                logger.error("テナントレベルでSMTP認証（基本認証）が無効化されています")
                logger.error("メール送信には代替手段（Graph API、Outlook COM等）が必要です")
                logger.error("詳細: https://aka.ms/smtp_auth_disabled")
            else:
                logger.error(f"SMTP認証エラー: {e}")
            raise
        except Exception as e:
            logger.error(f"メール送信エラー: {e}")
            logger.error(f"エラータイプ: {type(e).__name__}")
            raise

