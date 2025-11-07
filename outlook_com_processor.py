"""
Outlook COMオブジェクトを使用したメール処理
Windows限定、基本認証不要
"""
import win32com.client
from datetime import datetime, date, timedelta
from config import Config
from utils import get_logger

logger = get_logger(__name__)

class OutlookMailProcessor:
    """Outlook COMオブジェクトを使用したメール処理"""
    
    # Outlookフォルダタイプの定数
    # https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
    FOLDER_TYPE_MAIL = 0  # olFolderInbox など
    FOLDER_TYPE_CALENDAR = 9  # olFolderCalendar
    FOLDER_TYPE_CONTACTS = 10  # olFolderContacts
    FOLDER_TYPE_TASKS = 13  # olFolderTasks
    FOLDER_TYPE_NOTES = 12  # olFolderNotes
    FOLDER_TYPE_JOURNAL = 11  # olFolderJournal
    
    def __init__(self):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.inbox = self.namespace.GetDefaultFolder(6)  # 6 = Inbox
            self.excluded_folders = set()  # エラーが発生したフォルダを記録
            logger.info("Outlookに接続成功")
        except Exception as e:
            logger.error(f"Outlook接続エラー: {e}")
            logger.error("Outlookがインストールされ、起動している必要があります")
            raise
    
    def _should_exclude_folder(self, folder):
        """フォルダを除外すべきかどうかを判定"""
        try:
            folder_name = folder.Name
            
            # 既にエラーが発生したフォルダ
            if folder_name in self.excluded_folders:
                return True
            
            # 特定のフォルダを除外（同期の問題など、通常不要なフォルダ）
            exclude_by_name = [
                '同期の問題',
                'Sync Issues',
                'RSS フィード',
                'RSS Feeds',
                'Suggested Contacts',
                '提案された連絡先'
            ]
            for exclude_name in exclude_by_name:
                if exclude_name.lower() in folder_name.lower():
                    logger.debug(f"除外フォルダ（システムフォルダ）: {folder_name}")
                    return True
            
            # 設定で除外されたフォルダ名（部分一致）
            for exclude_pattern in Config.OUTLOOK_EXCLUDE_FOLDERS:
                if exclude_pattern.strip() and exclude_pattern.strip() in folder_name:
                    logger.debug(f"除外フォルダ（設定）: {folder_name}")
                    return True
            
            # フォルダタイプで除外判定
            try:
                # DefaultItemType で判定
                # 0 = olMailItem (メール)
                # 1 = olAppointmentItem (予定表)
                # 2 = olContactItem (連絡先)
                # 3 = olTaskItem (タスク)
                # 4 = olJournalItem (履歴)
                # 5 = olNoteItem (メモ)
                
                if hasattr(folder, 'DefaultItemType'):
                    item_type = folder.DefaultItemType
                    
                    # カレンダー（予定表）
                    if item_type == 1 and Config.OUTLOOK_EXCLUDE_CALENDAR:
                        logger.debug(f"除外フォルダ（カレンダー）: {folder_name}")
                        return True
                    
                    # 連絡先
                    if item_type == 2 and Config.OUTLOOK_EXCLUDE_CONTACTS:
                        logger.debug(f"除外フォルダ（連絡先）: {folder_name}")
                        return True
                    
                    # タスク
                    if item_type == 3 and Config.OUTLOOK_EXCLUDE_TASKS:
                        logger.debug(f"除外フォルダ（タスク）: {folder_name}")
                        return True
                    
                    # 履歴
                    if item_type == 4 and Config.OUTLOOK_EXCLUDE_JOURNAL:
                        logger.debug(f"除外フォルダ（ジャーナル）: {folder_name}")
                        return True
                    
                    # メモ
                    if item_type == 5 and Config.OUTLOOK_EXCLUDE_NOTES:
                        logger.debug(f"除外フォルダ（メモ）: {folder_name}")
                        return True
            except Exception as e:
                logger.debug(f"フォルダタイプ判定エラー（{folder_name}）: {e}")
            
            # フォルダ名でカレンダーを判定（バックアップ）
            if Config.OUTLOOK_EXCLUDE_CALENDAR:
                calendar_keywords = ['calendar', 'カレンダー', '予定表', '会議室']
                if any(keyword.lower() in folder_name.lower() for keyword in calendar_keywords):
                    logger.debug(f"除外フォルダ（カレンダー名）: {folder_name}")
                    return True
            
            return False
            
        except Exception as e:
            logger.warning(f"フォルダ除外判定エラー: {e}")
            # エラーが発生したフォルダは除外リストに追加
            try:
                self.excluded_folders.add(folder.Name)
            except:
                pass
            return True
    
    def _get_all_folders(self, parent_folder, folder_list):
        """フォルダを再帰的に取得（除外フォルダはスキップ）"""
        try:
            for folder in parent_folder.Folders:
                try:
                    # 除外判定
                    if self._should_exclude_folder(folder):
                        continue
                    
                    folder_list.append(folder)
                    
                    # サブフォルダも再帰的に取得
                    if folder.Folders.Count > 0:
                        self._get_all_folders(folder, folder_list)
                        
                except Exception as e:
                    # 個別のフォルダでエラーが発生した場合
                    try:
                        folder_name = folder.Name
                        logger.warning(f"フォルダ '{folder_name}' の処理でエラー: {e}")
                        # エラーが発生したフォルダを除外リストに追加
                        self.excluded_folders.add(folder_name)
                    except:
                        logger.warning(f"フォルダ処理でエラー: {e}")
                    continue
                    
        except Exception as e:
            logger.warning(f"フォルダ取得エラー: {e}")
    
    def _try_create_search_folder(self):
        """未読メール検索フォルダを自動作成する（現在は未サポート）"""
        logger.debug("検索フォルダの自動作成は現在サポートされていません")
        return None
    
    def _get_messages_from_search_folder(self, since_time):
        """検索フォルダから未読メールを取得（高速）"""
        try:
            logger.info("【最適化モード】未読メールを高速取得中...")
            
            search_folder = None
            
            # 方法1: すべてのGetDefaultFolderを試して検索フォルダを探す
            try:
                logger.debug("方法1: GetDefaultFolderで検索フォルダを探索中...")
                
                # Outlookのフォルダ定数を試す（1-20までの範囲）
                # 検索フォルダーは環境によって番号が異なる可能性がある
                for folder_num in range(1, 21):
                    try:
                        test_folder = self.namespace.GetDefaultFolder(folder_num)
                        folder_name = test_folder.Name
                        logger.debug(f"GetDefaultFolder({folder_num}): '{folder_name}'")
                        
                        # 「検索フォルダ」「Search Folders」という名前、または配下に検索フォルダがある
                        if any(keyword in folder_name for keyword in ["検索", "Search"]):
                            logger.info(f"✓ 検索関連フォルダ発見: GetDefaultFolder({folder_num}) = '{folder_name}'")
                            
                            # 配下のフォルダを確認
                            for sub_folder in test_folder.Folders:
                                logger.debug(f"  - {sub_folder.Name}")
                                
                                # 未読メールの検索フォルダを探す
                                if any(keyword in sub_folder.Name for keyword in ["未読", "Unread"]):
                                    search_folder = sub_folder
                                    logger.info(f"✓✓ 未読メール検索フォルダを発見: '{sub_folder.Name}'")
                                    break
                            
                            if search_folder:
                                break
                        
                        # 配下に「未読」フォルダがあるか確認
                        if not search_folder and test_folder.Folders.Count > 0:
                            for sub_folder in test_folder.Folders:
                                if any(keyword in sub_folder.Name for keyword in ["未読", "Unread"]) and "Mail" in sub_folder.Name:
                                    logger.info(f"✓ GetDefaultFolder({folder_num})配下に未読フォルダ発見: '{sub_folder.Name}'")
                                    search_folder = sub_folder
                                    break
                        
                        if search_folder:
                            break
                            
                    except Exception as e:
                        logger.debug(f"GetDefaultFolder({folder_num}): エラー - {e}")
                        continue
                
                if not search_folder:
                    logger.debug("方法1: すべてのGetDefaultFolderを試しましたが、検索フォルダが見つかりませんでした")
                    
            except Exception as e:
                logger.debug(f"方法1失敗: {type(e).__name__}: {e}")
            
            # 方法2: Storeから検索フォルダを探す
            if not search_folder:
                try:
                    logger.debug("方法2: Storeから検索フォルダを探索...")
                    default_store = self.namespace.GetDefaultFolder(6).Store
                    logger.debug(f"デフォルトストア: {default_store.DisplayName}")
                    
                    # Storeの全フォルダを再帰的に探索
                    root_folder = default_store.GetRootFolder()
                    logger.debug(f"ルートフォルダ: {root_folder.Name}")
                    
                    # すべてのフォルダをチェック
                    def find_search_folder(parent_folder, depth=0):
                        if depth > 3:  # 深さ制限
                            return None
                        
                        try:
                            for folder in parent_folder.Folders:
                                folder_name = folder.Name
                                logger.debug(f"{'  ' * depth}[深さ{depth}] {folder_name} (配下: {folder.Folders.Count if hasattr(folder, 'Folders') else 0}個)")
                                
                                # 検索フォルダコンテナまたは未読フォルダを探す
                                if folder_name in ["検索フォルダー", "Search Folders", "検索フォルダ"]:
                                    logger.info(f"✓✓ 検索フォルダコンテナ発見: '{folder_name}'")
                                    # 配下の未読フォルダを探す
                                    for sf in folder.Folders:
                                        logger.info(f"  {'  ' * depth}    - {sf.Name}")
                                        if any(kw in sf.Name for kw in ["未読", "Unread"]):
                                            logger.info(f"✓✓✓ 未読メール検索フォルダを発見: '{sf.Name}'")
                                            return sf
                                
                                # 再帰的に探索
                                result = find_search_folder(folder, depth + 1)
                                if result:
                                    return result
                        except Exception as e:
                            logger.debug(f"フォルダ探索エラー: {e}")
                        return None
                    
                    logger.debug("フォルダツリーを再帰的に探索中:")
                    search_folder = find_search_folder(root_folder)
                    
                    if search_folder:
                        logger.info(f"✓ 方法2で検索フォルダを発見: '{search_folder.Name}'")
                except Exception as e:
                    logger.debug(f"方法2失敗: {e}")
            
            # 検索フォルダが見つかった場合
            if search_folder:
                logger.info(f"検索フォルダ '{search_folder.Name}' からメールを取得します")
                return self._get_messages_from_specific_folder(search_folder, since_time)
            
            # 検索フォルダが見つからない場合
            if not search_folder:
                logger.debug("検索フォルダが見つかりませんでした")
                search_folder = self._try_create_search_folder()
            
            # 検索フォルダが見つかった/作成できた場合
            if search_folder:
                logger.info(f"✓✓ 検索フォルダ '{search_folder.Name}' を使用します")
                return self._get_messages_from_specific_folder(search_folder, since_time)
            
            # それでも見つからない場合はフォールバック
            logger.info("検索フォルダが利用できません → 全フォルダスキャンモードで処理を継続します（Restrictフィルタ使用で最適化済み）")
            
            # Noneを返して通常のフォルダスキャンにフォールバック
            return None
            
        except Exception as e:
            logger.error(f"検索フォルダ/受信トレイからのメール取得エラー: {e}", exc_info=True)
            return None
    
    def _get_messages_from_specific_folder(self, folder, since_time):
        """特定のフォルダ（検索フォルダ等）からメールを取得"""
        messages = []
        
        try:
            items = folder.Items
            
            # 日付フィルタを適用
            if since_time:
                date_filter = f"[ReceivedTime] >= '{since_time.strftime('%m/%d/%Y %H:%M')}'"
                try:
                    items = items.Restrict(date_filter)
                    logger.info(f"日付フィルタ適用: {items.Count}件")
                except Exception as e:
                    logger.warning(f"日付フィルタの適用に失敗: {e}")
            
            if items.Count == 0:
                logger.info("条件に一致するメールがありません")
                return []
            
            logger.info(f"検索フォルダ内のメール数: {items.Count}件")
            
            # ソート
            try:
                items.Sort("[ReceivedTime]", True)
            except:
                pass
            
            # メール情報を取得
            for item in items:
                try:
                    if not hasattr(item, 'Subject') or not hasattr(item, 'UnRead'):
                        continue
                    
                    subject = item.Subject or ""
                    body = item.Body or ""
                    sender = item.SenderName or ""
                    sender_email = item.SenderEmailAddress or ""
                    received = item.ReceivedTime
                    
                    to_field = ""
                    try:
                        to_field = item.To or ""
                    except:
                        pass
                    
                    # CC フィールドを取得
                    cc_field = ""
                    try:
                        cc_field = item.CC or ""
                    except:
                        pass
                    
                    # 元のフォルダ名を取得
                    folder_name = "不明"
                    try:
                        if hasattr(item, 'Parent') and item.Parent:
                            folder_name = item.Parent.Name
                    except:
                        pass
                    
                    # 日付を文字列に変換
                    if isinstance(received, datetime):
                        if hasattr(received, 'tzinfo') and received.tzinfo is not None:
                            received = received.replace(tzinfo=None)
                        date_str = received.strftime('%Y-%m-%d %H:%M:%S')
                    else:
                        date_str = str(received)
                    
                    messages.append({
                        'id': item.EntryID,
                        'subject': subject,
                        'from': f"{sender} <{sender_email}>",
                        'to': to_field,
                        'cc': cc_field,
                        'body': body,
                        'date': date_str,
                        'unread': item.UnRead,
                        'folder': folder_name
                    })
                except Exception as e:
                    logger.debug(f"メール情報の取得でエラー: {e}")
                    continue
            
            logger.info(f"フォルダ '{folder.Name}' から {len(messages)}件のメールを取得しました")
            return messages
            
        except Exception as e:
            logger.error(f"フォルダ '{folder.Name}' からのメール取得エラー: {e}", exc_info=True)
            return []
    
    def _get_messages_from_inbox_optimized(self, inbox, since_time):
        """受信トレイから最適化して未読メールを取得（検索フォルダの代替）"""
        messages = []
        
        try:
            logger.info("受信トレイから最適化取得を開始...")
            items = inbox.Items
            
            # 未読フィルタをRestrictで適用
            try:
                items = items.Restrict("[UnRead] = True")
                logger.info(f"未読フィルタ適用後: {items.Count}件")
            except Exception as e:
                logger.warning(f"未読フィルタの適用エラー: {e}")
                return None
            
            # 日付フィルタを適用
            if since_time:
                date_filter = f"[ReceivedTime] >= '{since_time.strftime('%m/%d/%Y %H:%M')}'"
                try:
                    items = items.Restrict(date_filter)
                    logger.info(f"日付フィルタ適用後: {items.Count}件")
                except Exception as e:
                    logger.warning(f"日付フィルタの適用エラー: {e}")
            
            if items.Count == 0:
                logger.info("条件に一致するメールがありません")
                return []
            
            # ソート
            try:
                items.Sort("[ReceivedTime]", True)  # 降順
            except:
                pass
            
            # メールを取得
            for item in items:
                try:
                    if not hasattr(item, 'Subject') or not hasattr(item, 'UnRead'):
                        continue
                    
                    subject = item.Subject or ""
                    body = item.Body or ""
                    sender = item.SenderName or ""
                    sender_email = item.SenderEmailAddress or ""
                    received = item.ReceivedTime
                    
                    to_field = ""
                    try:
                        to_field = item.To or ""
                    except:
                        pass
                    
                    # CC フィールドを取得
                    cc_field = ""
                    try:
                        cc_field = item.CC or ""
                    except:
                        pass
                    
                    # 日付を文字列に変換
                    if isinstance(received, datetime):
                        if hasattr(received, 'tzinfo') and received.tzinfo is not None:
                            received = received.replace(tzinfo=None)
                        date_str = received.strftime('%Y-%m-%d %H:%M:%S')
                    else:
                        date_str = str(received)
                    
                    messages.append({
                        'id': item.EntryID,
                        'subject': subject,
                        'from': f"{sender} <{sender_email}>",
                        'to': to_field,
                        'cc': cc_field,
                        'body': body,
                        'date': date_str,
                        'unread': item.UnRead,
                        'folder': inbox.Name
                    })
                except Exception as e:
                    logger.debug(f"メール情報の取得でエラー: {e}")
                    continue
            
            logger.info(f"受信トレイから {len(messages)}件のメールを取得しました")
            return messages
            
        except Exception as e:
            logger.error(f"受信トレイからの最適化取得エラー: {e}")
            return None
    
    def _get_messages_from_folder(self, folder, unread_only=True, since_time=None):
        """指定フォルダからメールを取得"""
        messages = []
        
        try:
            # フォルダのアイテムを取得
            items = folder.Items
            item_count = items.Count
            
            # アイテムが空の場合はスキップ
            if item_count == 0:
                return messages
            
            # 大量のアイテムがあるフォルダは警告（処理に時間がかかる可能性）
            folder_name = folder.Name
            if item_count > 1000:
                logger.warning(f"フォルダ '{folder_name}' には {item_count} 件のアイテムがあります。処理に時間がかかる可能性があります。")
            elif item_count > 100:
                logger.debug(f"フォルダ '{folder_name}' のアイテム数: {item_count}件")
            
            # 未読フィルタをRestrictで適用（高速化）
            if unread_only:
                try:
                    items = items.Restrict("[UnRead] = True")
                    item_count = items.Count
                    if item_count == 0:
                        return messages
                    logger.debug(f"フォルダ '{folder_name}': 未読メール {item_count}件")
                except Exception as restrict_error:
                    logger.debug(f"フォルダ '{folder_name}' の未読フィルタエラー: {restrict_error}")
                    # フィルタが失敗した場合は後でチェック
            
            # 並び替え（受信日時の降順）
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception as sort_error:
                logger.debug(f"フォルダ '{folder_name}' のソートエラー: {sort_error}")
                pass  # 一部のフォルダはソートできない場合がある
            
            processed_count = 0
            for item in items:
                processed_count += 1
                # 進捗ログ（大きいフォルダの場合）
                if item_count > 500 and processed_count % 100 == 0:
                    logger.debug(f"フォルダ '{folder_name}': {processed_count}/{item_count} 件処理中...")
                try:
                    # メールアイテムのみ処理（予定表や連絡先は除外）
                    if not hasattr(item, 'Subject') or not hasattr(item, 'UnRead'):
                        continue
                    
                    # 未読フィルタ（Restrictが効かなかった場合のバックアップ）
                    if unread_only and hasattr(item, 'UnRead') and item.UnRead == False:
                        continue
                    
                    # 日付フィルタ
                    if since_time:
                        received_time = item.ReceivedTime
                        if isinstance(received_time, str):
                            received_time = datetime.strptime(received_time, "%Y-%m-%d %H:%M:%S")
                        elif isinstance(received_time, datetime):
                            # タイムゾーン情報を削除して比較
                            if hasattr(received_time, 'tzinfo') and received_time.tzinfo is not None:
                                received_time = received_time.replace(tzinfo=None)
                        if received_time < since_time:
                            continue
                    
                    # メール情報を取得
                    subject = item.Subject or ""
                    body = item.Body or ""
                    sender = item.SenderName or ""
                    sender_email = item.SenderEmailAddress or ""
                    received = item.ReceivedTime
                    
                    # To フィールドを取得（ルールベース判定用）
                    to_field = ""
                    try:
                        to_field = item.To or ""
                    except:
                        pass
                    
                    # CC フィールドを取得
                    cc_field = ""
                    try:
                        cc_field = item.CC or ""
                    except:
                        pass
                    
                    # 日付を文字列に変換
                    if isinstance(received, datetime):
                        # タイムゾーン情報を削除
                        if hasattr(received, 'tzinfo') and received.tzinfo is not None:
                            received = received.replace(tzinfo=None)
                        date_str = received.strftime('%Y-%m-%d %H:%M:%S')
                    else:
                        date_str = str(received)
                    
                    messages.append({
                        'id': item.EntryID,  # Outlookの一意ID
                        'subject': subject,
                        'from': f"{sender} <{sender_email}>",
                        'to': to_field,  # To フィールドを追加
                        'cc': cc_field,  # CC フィールドを追加
                        'body': body,
                        'date': date_str,
                        'unread': item.UnRead,
                        'folder': folder.Name  # フォルダ名を追加
                    })
                except Exception as e:
                    logger.debug(f"メール情報の取得でエラー: {e}")
                    continue
                    
        except Exception as e:
            try:
                folder_name = folder.Name
                logger.warning(f"フォルダ '{folder_name}' のメール取得エラー: {e}")
                # エラーが発生したフォルダを除外リストに追加
                self.excluded_folders.add(folder_name)
            except:
                logger.warning(f"メール取得エラー: {e}")
        
        return messages
    
    def get_messages(self, unread_only=True, since_time=None):
        """全フォルダからメールを取得"""
        all_messages = []
        
        try:
            # 起動時刻から24時間前をデフォルトとして設定
            if since_time is None:
                # 現在時刻から24時間前
                since_time = datetime.now() - timedelta(hours=24)
                logger.info(f"デフォルト時間範囲: 過去24時間")
            
            # since_timeをタイムゾーンなし（naive）にする
            if hasattr(since_time, 'tzinfo') and since_time.tzinfo is not None:
                since_time = since_time.replace(tzinfo=None)
            
            logger.info(f"メール取得開始: {since_time.strftime('%Y-%m-%d %H:%M:%S')}以降（現在: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}）")
            
            # === 方法1: 検索フォルダを使用（高速） ===
            if unread_only:
                try:
                    search_result = self._get_messages_from_search_folder(since_time)
                    if search_result is not None:  # 検索フォルダが見つかった場合
                        logger.info(f"検索フォルダから合計 {len(search_result)}件のメールを取得しました")
                        return search_result
                    else:
                        logger.info("検索フォルダが見つかりませんでした。通常のフォルダスキャンを実行します。")
                except Exception as e:
                    logger.warning(f"検索フォルダの使用に失敗: {e}")
                    logger.info("通常のフォルダスキャンにフォールバックします。")
            
            # すべてのフォルダを取得
            all_folders = []
            
            # すべてのメールアカウントをループ
            try:
                folder_count = self.namespace.Folders.Count
                logger.info(f"検出されたメールアカウント数: {folder_count}件")
                
                for i in range(1, folder_count + 1):
                    try:
                        root_folder = self.namespace.Folders.Item(i)
                        logger.info(f"アカウント '{root_folder.Name}' のフォルダを取得中...")
                        
                        # このアカウント配下のすべてのフォルダを再帰的に取得
                        self._get_all_folders(root_folder, all_folders)
                    except Exception as e:
                        logger.warning(f"アカウント {i} の処理でエラー: {e}")
                        continue
            except Exception as e:
                logger.warning(f"フォルダ一覧の取得でエラー: {e}")
                # フォールバック: 最初のアカウントのみ取得
                try:
                    root_folder = self.namespace.Folders.Item(1)
                    logger.info(f"フォールバック: アカウント '{root_folder.Name}' のフォルダを取得中...")
                    self._get_all_folders(root_folder, all_folders)
                except Exception as e2:
                    logger.error(f"フォールバック処理もエラー: {e2}")
            
            # 受信トレイも含める（重複チェック）
            inbox_included = any(folder.Name == self.inbox.Name for folder in all_folders)
            if not inbox_included:
                all_folders.insert(0, self.inbox)
                logger.debug("受信トレイを追加しました")
            
            logger.info(f"検索対象フォルダ数: {len(all_folders)}件")
            
            # 検索対象フォルダの一覧を表示
            if all_folders:
                logger.info("検索対象フォルダ一覧:")
                for idx, folder in enumerate(all_folders[:10], 1):  # 最初の10件を表示
                    try:
                        logger.info(f"  {idx}. {folder.Name}")
                    except:
                        logger.info(f"  {idx}. (名前取得エラー)")
                if len(all_folders) > 10:
                    logger.info(f"  ... 他 {len(all_folders) - 10}件")
            
            if self.excluded_folders:
                logger.info(f"除外されたフォルダ: {len(self.excluded_folders)}件")
                for excluded in list(self.excluded_folders)[:5]:  # 最初の5件を表示
                    logger.debug(f"  - {excluded}")
                if len(self.excluded_folders) > 5:
                    logger.debug(f"  ... 他 {len(self.excluded_folders) - 5}件")
            
            # 各フォルダからメールを取得
            logger.info(f"フォルダスキャン開始（合計 {len(all_folders)}件）")
            for idx, folder in enumerate(all_folders, 1):
                try:
                    folder_name = folder.Name
                    # 進捗ログ（10件ごと、または重要なフォルダ）
                    if idx % 10 == 0 or idx <= 20 or idx == len(all_folders):
                        logger.info(f"進捗: {idx}/{len(all_folders)} - 処理中: {folder_name}")
                    
                    folder_messages = self._get_messages_from_folder(folder, unread_only, since_time)
                    if folder_messages:
                        logger.info(f"フォルダ '{folder_name}' から {len(folder_messages)}件のメールを取得")
                        all_messages.extend(folder_messages)
                except Exception as e:
                    try:
                        folder_name = folder.Name
                        logger.warning(f"フォルダ '{folder_name}' の処理でエラー: {e}")
                        # エラーが発生したフォルダを除外リストに追加
                        self.excluded_folders.add(folder_name)
                    except:
                        logger.warning(f"フォルダ処理でエラー ({idx}/{len(all_folders)}): {e}")
                    continue
            
            logger.info("フォルダスキャン完了")
            
            logger.info(f"合計 {len(all_messages)}件のメールを取得しました")
                    
        except Exception as e:
            logger.error(f"メール取得エラー: {e}", exc_info=True)
            # エラーが発生しても空のリストを返す
            if all_messages is None:
                all_messages = []
        
        # 念のため、all_messagesがNoneの場合は空リストを返す
        if all_messages is None:
            logger.warning("all_messagesがNoneです。空のリストを返します。")
            return []
        
        return all_messages
    
    def is_target_mail(self, message):
        """ターゲットメールかどうかを判定"""
        body = message.get('body', '').lower()
        return any(keyword.lower() in body for keyword in Config.KEYWORDS)
    
    def mark_as_read(self, message_id):
        """メールを既読にする"""
        try:
            # EntryIDでメールを取得
            item = self.namespace.GetItemFromID(message_id)
            item.UnRead = False
            item.Save()
            logger.info(f"メールを既読にしました: {message_id}")
        except Exception as e:
            logger.error(f"既読処理エラー: {e}")
    
    def mark_as_unread(self, message_id):
        """メールを未読にする"""
        try:
            # EntryIDでメールを取得
            item = self.namespace.GetItemFromID(message_id)
            item.UnRead = True
            item.Save()
            logger.info(f"メールを未読にしました: {message_id}")
        except Exception as e:
            logger.error(f"未読処理エラー: {e}")
            raise
    
    def send_email(self, subject, body, to_address):
        """メールを送信または下書きに保存"""
        try:
            mail = self.outlook.CreateItem(0)  # 0 = MailItem
            mail.Subject = subject
            mail.Body = body
            mail.To = to_address
            
            # 設定に応じて送信または下書き保存
            if Config.SAVE_AS_DRAFT:
                # 下書きに保存
                mail.Save()
                logger.info(f"メールを下書きに保存: {subject}")
            else:
                # 送信
                mail.Send()
                logger.info(f"メール送信成功: {subject}")
        except Exception as e:
            logger.error(f"メール送信/保存エラー: {e}")
            raise
    
    def get_sent_messages(self, limit=5):
        """送信済みメールを取得（文体学習用）"""
        messages = []
        
        try:
            # 送信済みアイテムフォルダを取得
            sent_folder = self.namespace.GetDefaultFolder(5)  # 5 = SentMail
            logger.debug(f"送信済みフォルダ: '{sent_folder.Name}'")
            
            items = sent_folder.Items
            
            # 最新順にソート
            try:
                items.Sort("[SentOn]", True)  # 降順
            except:
                pass
            
            # 最新のlimit件を取得
            count = 0
            for item in items:
                if count >= limit:
                    break
                
                try:
                    if not hasattr(item, 'Subject') or not hasattr(item, 'Body'):
                        continue
                    
                    subject = item.Subject or ""
                    body = item.Body or ""
                    
                    # 返信・転送メールは除外（オリジナルのメールのみ）
                    if subject.startswith('Re:') or subject.startswith('FW:') or subject.startswith('Fw:'):
                        continue
                    
                    messages.append({
                        'subject': subject,
                        'body': body,
                        'date': str(item.SentOn) if hasattr(item, 'SentOn') else ''
                    })
                    
                    count += 1
                    
                except Exception as e:
                    logger.debug(f"送信済みメール取得エラー: {e}")
                    continue
            
            logger.debug(f"送信済みメール {len(messages)}件を取得しました")
            return messages
            
        except Exception as e:
            logger.error(f"送信済みメール取得エラー: {e}")
            return []
    
    def disconnect(self):
        """接続を閉じる（COMオブジェクトの場合は不要だが、一貫性のため）"""
        try:
            # COMオブジェクトは自動的に解放される
            self.outlook = None
            self.namespace = None
            self.inbox = None
            logger.info("Outlook接続を閉じました")
        except Exception as e:
            logger.error(f"切断エラー: {e}")

