from datetime import datetime, timedelta
import pytz
import sys
import os
from pathlib import Path
from config import Config
from ai_processor import AIProcessor
from storage_manager import StorageManager
from utils import get_logger, get_jst_now, get_time_range_for_daily_report

logger = get_logger(__name__)

# フローCの実行日を記録するファイル
FLOW_C_LAST_RUN_FILE = Path('logs') / 'flow_c_last_run.txt'

def get_email_processing_type(message):
    """メールの処理タイプを判定
    
    戻り値:
        -2: 除外メール（既読にする、AI処理なし）- セールスメールなど
        -1: 除外メール（未読維持、AI処理なし）- Teams通知など
        0: 既読にする（AI処理なし）
        1: 要約のみ（未読維持）
        2: 要約+返信下書き（未読維持）
    """
    # 最優先: 除外送信元チェック（すべてのモードで有効）
    from_address = message.get('from', '').lower()
    
    # パターン1: 除外して既読にする（セールスメールなど）
    for exclude_domain in Config.EXCLUDE_AND_READ_DOMAINS:
        if exclude_domain.strip() and exclude_domain.strip().lower() in from_address:
            logger.info(f"[判定] 除外ドメイン ({exclude_domain.strip()}) から送信 → AI処理スキップ → 既読にする")
            return -2
    
    # パターン2: 除外して未読維持（Teams通知など）
    for exclude_domain in Config.EXCLUDE_SENDER_DOMAINS:
        if exclude_domain.strip() and exclude_domain.strip().lower() in from_address:
            logger.info(f"[判定] 除外ドメイン ({exclude_domain.strip()}) から送信 → AI処理スキップ → 未読維持")
            return -1
    
    mode = Config.MARK_AS_READ_AFTER_PROCESS
    
    # モード1: 全部未読のまま（要約+返信）
    if mode == 1:
        return 2
    
    # モード2: 全部既読にする
    if mode == 2:
        return 0
    
    # モード0: ルールベース判定
    if mode == 0:
        my_email = Config.EMAIL_ADDRESS.lower()
        
        to_field = message.get('to', '').lower()
        cc_field = message.get('cc', '').lower()
        subject = message.get('subject', '')
        body = message.get('body', '')
        
        # デバッグ: 実際の値を確認
        logger.debug(f"[判定開始] 件名: {subject[:50]}")
        logger.debug(f"  - 自分のアドレス: {my_email}")
        logger.debug(f"  - Toフィールド: {to_field[:100] if to_field else '(空)'}")
        logger.debug(f"  - CCフィールド: {cc_field[:100] if cc_field else '(空)'}")
        
        # ★重要★ Toフィールドが空の場合は、すべてのメールを自分宛として扱う
        if not to_field:
            logger.warning(f"[判定] Toフィールドが空です → 安全のため自分宛として扱います")
            logger.warning(f"  - 件名: {subject[:100]}")
            return 2  # 要約+返信
        
        # 最優先: グループメールアドレス宛のチェック（個人アドレスより先に判定）
        for group_email in Config.GROUP_EMAIL_ADDRESSES:
            if group_email.strip() and group_email.strip().lower() in to_field:
                # グループメール宛の場合、Toに個人アドレスも含まれていても名前チェックを優先
                if Config.MY_NAME and Config.MY_NAME in body:
                    logger.info(f"[判定] グループアドレス宛 ({group_email.strip()}) + 本文に名前あり → 要約のみ → 未読維持")
                    return 1
                else:
                    logger.info(f"[判定] グループアドレス宛 ({group_email.strip()}) だが本文に名前なし → 既読にする")
                    logger.info(f"  - 本文に '{Config.MY_NAME}' が見つかりませんでした")
                    return 0
        
        # パターン1: To フィールドに自分のメールアドレスが含まれている → 要約+返信
        # （グループメール宛でない場合のみ到達）
        if my_email and my_email in to_field:
            logger.info(f"[判定] To に自分のアドレスあり → 要約+返信下書き → 未読維持")
            return 2
        
        # パターン2: CCに自分のアドレス → 要約のみ
        if my_email and my_email in cc_field:
            logger.info(f"[判定] CC に自分のアドレスあり → 要約のみ → 未読維持")
            return 1
        
        # BCC対応: To/CCに自分のアドレスがないが、本文または件名に名前がある場合
        # （BCCや展開されたグループメール対応）
        if Config.MY_NAME:
            if Config.MY_NAME in body:
                logger.info(f"[判定] To/CCに自分のアドレスなし + 本文に名前あり → 要約のみ → 未読維持（BCC/展開メール対応）")
                return 1
            elif Config.MY_NAME in subject:
                logger.info(f"[判定] To/CCに自分のアドレスなし + 件名に名前あり → 要約のみ → 未読維持（BCC/展開メール対応）")
                logger.info(f"  - 件名: {subject[:100]}")
                return 1
        
        # それ以外は既読にする
        logger.info(f"[判定] 条件に該当せず → 既読にする")
        logger.info(f"  - To: {to_field[:100]}")
        return 0
    
    # デフォルト: 既読にする
    return 0

# メールプロセッサーの選択
if Config.MAIL_PROCESSOR_TYPE == 'outlook_com':
    try:
        from outlook_com_processor import OutlookMailProcessor as MailProcessor
        logger.info("Outlook COMオブジェクトを使用します (Windows専用)")
    except ImportError as e:
        logger.error(f"Outlook COMプロセッサーのインポートエラー: {e}")
        logger.error("pywin32がインストールされているか確認してください: pip install pywin32")
        sys.exit(1)
elif Config.MAIL_PROCESSOR_TYPE == 'graph':
    try:
        from graph_mail_processor import GraphMailProcessor as MailProcessor
        logger.info("Microsoft Graph APIを使用します")
    except ImportError as e:
        logger.error(f"Graph APIプロセッサーのインポートエラー: {e}")
        sys.exit(1)
elif Config.MAIL_PROCESSOR_TYPE == 'ews':
    try:
        from ews_mail_processor import EWSMailProcessor as MailProcessor
        logger.info("Exchange Web Services (EWS)を使用します (クロスプラットフォーム対応)")
    except ImportError as e:
        logger.error(f"EWSプロセッサーのインポートエラー: {e}")
        logger.error("exchangelibがインストールされているか確認してください: pip install exchangelib")
        sys.exit(1)
else:
    from mail_processor import MailProcessor
    logger.info("IMAP/SMTPを使用します")

def flow_a_process_mails():
    """フローA: メール処理、要約、下書き作成"""
    logger.info("=== フローA 開始 ===")
    
    mail_processor = MailProcessor()
    ai_processor = AIProcessor(mail_processor=mail_processor)
    storage_manager = StorageManager()
    
    try:
        # メール取得
        messages = mail_processor.get_messages(unread_only=True)
        if messages is None:
            logger.error("メール取得に失敗しました（Noneが返されました）")
            messages = []
        logger.info(f"未読メール数: {len(messages)}")
        
        for idx, message in enumerate(messages, 1):
            logger.info(f"=== メール処理 {idx}/{len(messages)} ===")
            
            subject = message.get('subject', '')
            body = message.get('body', '')
            from_address = message.get('from', '')
            message_id = message.get('id')
            folder_name = message.get('folder', '不明')
            
            logger.info(f"処理中: [{folder_name}] {subject}")
            
            # ★★★ 重要: メール処理タイプを判定 ★★★
            processing_type = get_email_processing_type(message)
            
            if processing_type == -2:
                # ========================================
                # 除外メール（既読）: AI処理なし＋既読にする（セールスメールなど）
                # ========================================
                logger.info("🗑️ [除外・既読] 除外ドメインから送信 → AI処理スキップ → 既読にする")
                mail_processor.mark_as_read(message_id)
                logger.info(f"✓ [除外・既読] [{folder_name}] {subject}")
            
            elif processing_type == -1:
                # ========================================
                # 除外メール（未読）: AI処理なし＋未読維持（Teams通知など）
                # ========================================
                logger.info("🚫 [除外・未読] 除外ドメインから送信 → AI処理スキップ → 未読維持")
                try:
                    mail_processor.mark_as_unread(message_id)
                    logger.info(f"✉️ [除外・未読維持] [{folder_name}] {subject}")
                except Exception as e:
                    logger.warning(f"未読に戻す処理でエラー: {e}")
                    logger.info(f"✉️ [除外・未読維持を試みました] [{folder_name}] {subject}")
            
            elif processing_type == 2:
                # ========================================
                # パターン1: 要約＋返信下書き＋Obsidian保存＋未読維持
                # ========================================
                logger.info("📧 [パターン1] Toに自分のアドレスあり → 要約＋返信下書き＋Obsidian保存を実行します")
                
                # 要約生成
                logger.info("AI処理ステップ1: 要約生成開始")
                summary = ai_processor.summarize_email(subject, body)
                if summary is None:
                    summary = "（要約生成に失敗しました。OpenAI APIのクォータを確認してください）"
                    logger.warning(f"要約生成に失敗: {subject}")
                logger.info("AI処理ステップ1: 要約生成終了")
                
                # 返信文生成
                logger.info("AI処理ステップ2: 返信文生成開始")
                reply = ai_processor.generate_reply(subject, body, from_address=from_address)
                if reply is None:
                    logger.warning(f"返信文生成に失敗: {subject}")
                logger.info("AI処理ステップ2: 返信文生成終了")
                
                # メール送信または下書き保存（返信が必要な場合のみ、かつ送信が有効な場合）
                if reply and Config.ENABLE_EMAIL_SENDING:
                    try:
                        mail_processor.send_email(
                            subject=f"Re: {subject}",
                            body=reply,
                            to_address=from_address
                        )
                        if Config.SAVE_AS_DRAFT:
                            logger.info(f"返信を下書きに保存しました: {subject}")
                        else:
                            logger.info(f"返信を送信しました: {subject}")
                    except Exception as e:
                        logger.warning(f"メール送信/保存に失敗しました: {e}")
                elif reply and not Config.ENABLE_EMAIL_SENDING:
                    logger.info(f"メール送信が無効化されています。返信文を生成しましたが送信しません。")
                
                # TODO抽出
                logger.info("AI処理ステップ3: TODO抽出開始")
                todos = ai_processor.extract_todos(body)
                if todos is None:
                    todos = "（TODO抽出に失敗しました）"
                    logger.warning(f"TODO抽出に失敗: {subject}")
                logger.info("AI処理ステップ3: TODO抽出終了")
                
                # Obsidianに保存
                logger.info("Obsidian保存開始")
                from utils import format_markdown_email_summary, extract_sender_name
                content = format_markdown_email_summary(
                    subject=subject,
                    from_address=from_address,
                    summary=summary,
                    todos=todos,
                    folder=folder_name,
                    message=message
                )
                sender_name = extract_sender_name(from_address)
                storage_manager.save_to_obsidian(content, subject[:50], sender_name=sender_name)
                logger.info("Obsidian保存完了")
                
                # 自分宛のメールは未読に戻す（明示的に未読にする）
                try:
                    mail_processor.mark_as_unread(message_id)
                    logger.info(f"✉️ [未読に戻しました] [{folder_name}] {subject}")
                except Exception as e:
                    logger.warning(f"未読に戻す処理でエラー: {e}")
                    logger.info(f"✉️ [未読のまま維持を試みました] [{folder_name}] {subject}")
            
            elif processing_type == 1:
                # ========================================
                # パターン2: 要約のみ＋Obsidian保存＋未読維持（返信なし）
                # ========================================
                logger.info("📬 [パターン2] CC/グループ宛+名前あり → 要約のみ＋Obsidian保存を実行します（返信なし）")
                
                # 要約生成
                logger.info("AI処理ステップ1: 要約生成開始")
                summary = ai_processor.summarize_email(subject, body)
                if summary is None:
                    summary = "（要約生成に失敗しました。OpenAI APIのクォータを確認してください）"
                    logger.warning(f"要約生成に失敗: {subject}")
                logger.info("AI処理ステップ1: 要約生成終了")
                
                # TODO抽出
                logger.info("AI処理ステップ2: TODO抽出開始")
                todos = ai_processor.extract_todos(body)
                if todos is None:
                    todos = "（TODO抽出に失敗しました）"
                    logger.warning(f"TODO抽出に失敗: {subject}")
                logger.info("AI処理ステップ2: TODO抽出終了")
                
                # Obsidianに保存
                logger.info("Obsidian保存開始")
                from utils import format_markdown_email_summary, extract_sender_name
                content = format_markdown_email_summary(
                    subject=subject,
                    from_address=from_address,
                    summary=summary,
                    todos=todos,
                    folder=folder_name,
                    message=message
                )
                sender_name = extract_sender_name(from_address)
                storage_manager.save_to_obsidian(content, subject[:50], sender_name=sender_name)
                logger.info("Obsidian保存完了")
                
                # 未読に戻す（返信は生成しない）
                try:
                    mail_processor.mark_as_unread(message_id)
                    logger.info(f"✉️ [要約のみ完了・未読維持] [{folder_name}] {subject}")
                except Exception as e:
                    logger.warning(f"未読に戻す処理でエラー: {e}")
                    logger.info(f"✉️ [要約のみ完了・未読維持を試みました] [{folder_name}] {subject}")
                
            else:
                # ========================================
                # パターン0: 既読のみ（AI処理なし）
                # ========================================
                logger.info("📮 [パターン0] 条件に該当せず → 既読にします（AI処理スキップ）")
                mail_processor.mark_as_read(message_id)
                logger.info(f"✓ [既読] [{folder_name}] {subject}")
            
            logger.info(f"=== メール処理完了 {idx}/{len(messages)} ===")
    finally:
        mail_processor.disconnect()
    
    logger.info("=== フローA 終了 ===")

def flow_b_check_monitoring_mails():
    """フローB: 運用定期メールの監視"""
    logger.info("=== フローB 開始 ===")
    
    mail_processor = MailProcessor()
    ai_processor = AIProcessor(mail_processor=mail_processor)
    storage_manager = StorageManager()
    
    try:
        messages = mail_processor.get_messages(unread_only=True)
        if messages is None:
            logger.error("メール取得に失敗しました（Noneが返されました）")
            messages = []
        
        for message in messages:
            subject = message.get('subject', '')
            body = message.get('body', '')
            
            status = ai_processor.check_abnormal(subject, body)
            
            if status == 'normal':
                # 既読/未読を判定
                processing_type = get_email_processing_type(message)
                if processing_type == -2:
                    # 除外メール（既読）
                    mail_processor.mark_as_read(message['id'])
                    logger.info(f"[除外・既読] 正常メール: {subject}")
                elif processing_type == -1:
                    # 除外メール（未読）
                    mail_processor.mark_as_unread(message['id'])
                    logger.info(f"[除外・未読] 正常メール: {subject}")
                elif processing_type == 0:
                    mail_processor.mark_as_read(message['id'])
                    logger.info(f"[既読] 正常メール: {subject}")
                else:
                    logger.info(f"[未読] 正常メール: {subject}")
            elif status == 'abnormal':
                # Slack通知
                storage_manager.send_slack_notification(
                    f"[異常検知]\n件名: {subject}"
                )
    finally:
        mail_processor.disconnect()
    
    logger.info("=== フローB 終了 ===")

def should_run_flow_c():
    """フローCを実行すべきかどうかを判定
    
    戻り値:
        True: 実行すべき（今日まだ実行していない）
        False: 実行不要（今日すでに実行済み）
    """
    today = datetime.now().strftime('%Y-%m-%d')
    
    # 記録ファイルが存在しない場合は実行する
    if not FLOW_C_LAST_RUN_FILE.exists():
        logger.info(f"[フローC判定] 実行記録なし → 実行します")
        return True
    
    try:
        # 最後に実行した日付を読み込む
        with open(FLOW_C_LAST_RUN_FILE, 'r', encoding='utf-8') as f:
            last_run_date = f.read().strip()
        
        if last_run_date == today:
            logger.info(f"[フローC判定] 本日すでに実行済み ({last_run_date}) → スキップします")
            return False
        else:
            logger.info(f"[フローC判定] 前回実行: {last_run_date} → 本日分を実行します")
            return True
    except Exception as e:
        logger.warning(f"[フローC判定] 記録ファイル読み込みエラー: {e} → 念のため実行します")
        return True

def record_flow_c_execution():
    """フローCの実行日を記録する"""
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        FLOW_C_LAST_RUN_FILE.parent.mkdir(parents=True, exist_ok=True)
        with open(FLOW_C_LAST_RUN_FILE, 'w', encoding='utf-8') as f:
            f.write(today)
        logger.info(f"[フローC記録] 実行日を記録しました: {today}")
    except Exception as e:
        logger.warning(f"[フローC記録] 記録ファイル書き込みエラー: {e}")

def flow_c_daily_report():
    """フローC: 日次レポート作成"""
    logger.info("=== フローC 開始 ===")
    
    mail_processor = MailProcessor()
    storage_manager = StorageManager()
    
    try:
        start_time, end_time = get_time_range_for_daily_report()
        
        messages = mail_processor.get_messages(
            unread_only=False,
            since_time=start_time
        )
        if messages is None:
            logger.error("メール取得に失敗しました（Noneが返されました）")
            messages = []
        
        report = f"""# 日次メールレポート

**集計期間**: {start_time.strftime('%Y-%m-%d %H:%M')} ～ {end_time.strftime('%Y-%m-%d %H:%M')}

## 統計
- 受信メール数: {len(messages)}

## 詳細
"""
        
        for message in messages[:20]:  # 最新20件
            subject = message.get('subject', '')
            from_address = message.get('from', '')
            report += f"- {subject} from {from_address}\n"
        
        storage_manager.save_to_obsidian(
            report,
            f"daily_report_{datetime.now().strftime('%Y%m%d')}",
            folder='Reports'
        )
        
        # 実行日を記録
        record_flow_c_execution()
    finally:
        mail_processor.disconnect()
    
    logger.info("=== フローC 終了 ===")

def main():
    """メイン処理"""
    jst_now = get_jst_now()
    logger.info(f"メール自動処理開始: {jst_now}")
    logger.info(f"ログレベル: {Config.LOG_LEVEL}")
    
    # 既読モードの表示
    mode = Config.MARK_AS_READ_AFTER_PROCESS
    logger.info(f"[設定確認] MARK_AS_READ_AFTER_PROCESS = {mode}")
    if mode == 0:
        logger.info("[既読モード] ルールベース（To/名前で判定）")
        if Config.MY_NAME:
            logger.info(f"  - 自分の名前: {Config.MY_NAME}")
        else:
            logger.info(f"  - 自分の名前: (未設定)")
        logger.info(f"  - 自分のアドレス: {Config.EMAIL_ADDRESS}")
    elif mode == 1:
        logger.info("[既読モード] すべて未読のまま維持")
    elif mode == 2:
        logger.info("[既読モード] すべて既読にする")
    else:
        logger.warning(f"[警告] 既読モード: 不明な値 ({mode})、デフォルト（未読）を使用")
    
    try:
        # フローA: メール処理
        flow_a_process_mails()
        
        # フローB: 監視メールチェック
        flow_b_check_monitoring_mails()
        
        # フローC: 日次レポート（1日1回、8時以降に実行）
        if jst_now.hour >= 8 and should_run_flow_c():
            flow_c_daily_report()
        
    except Exception as e:
        logger.error(f"エラー発生: {e}", exc_info=True)
        # エラーはSlackに通知
        storage_manager = StorageManager()
        storage_manager.send_slack_notification(f"❌ メール自動処理エラー: {str(e)}")
    
    logger.info("メール自動処理終了")

if __name__ == "__main__":
    main()

