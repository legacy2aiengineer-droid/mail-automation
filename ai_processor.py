from openai import OpenAI
from config import Config
from utils import get_logger
import httpx
from pathlib import Path

logger = get_logger(__name__)

class AIProcessor:
    def __init__(self, mail_processor=None):
        # タイムアウト設定（60秒）
        timeout = httpx.Timeout(60.0, connect=10.0)
        self.client = OpenAI(
            api_key=Config.OPENAI_API_KEY,
            timeout=timeout
        )
        self.mail_processor = mail_processor
        self.writing_style_samples = []  # 文体サンプルを保存
        
        # 初期化時に送信済みメールから文体を学習
        if Config.LEARN_WRITING_STYLE and mail_processor:
            self._load_writing_style()
    
    def _load_writing_style(self):
        """送信済みメールから文体を学習"""
        try:
            logger.info("送信済みメールから文体を学習中...")
            
            # 送信済みメールを取得（設定値の件数）
            limit = Config.WRITING_STYLE_SAMPLE_COUNT
            sent_messages = self.mail_processor.get_sent_messages(limit=limit)
            
            if not sent_messages:
                logger.warning("送信済みメールが取得できませんでした")
                return
            
            logger.info(f"送信済みメール {len(sent_messages)}件を文体サンプルとして取得しました")
            
            for msg in sent_messages:
                body = msg.get('body', '')
                if body and len(body) > 50:  # 十分な長さのメールのみ
                    # メールの最初の500文字を保存
                    self.writing_style_samples.append(body[:500])
            
            logger.info(f"文体サンプル数: {len(self.writing_style_samples)}件")
            
        except Exception as e:
            logger.warning(f"文体学習エラー: {e}")
            # エラーが発生しても処理は継続
    
    def summarize_email(self, subject, body):
        """メール内容を要約"""
        logger.info(f"メール要約開始: {subject[:50]}...")
        prompt = f"""
以下のメールを要約してください。重要な情報を簡潔にまとめてください。

件名: {subject}

本文:
{body[:1000]}

要約:
"""
        try:
            response = self.client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "あなたはメール要約の専門家です。"},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=500
            )
            logger.info(f"メール要約完了: {subject[:50]}...")
            return response.choices[0].message.content
        except Exception as e:
            logger.error(f"メール要約エラー ({subject[:50]}...): {e}", exc_info=True)
            return None
    
    def _load_template(self, template_path):
        """テンプレートファイルを読み込む"""
        try:
            template_file = Path(template_path)
            if template_file.exists():
                with open(template_file, 'r', encoding='utf-8') as f:
                    return f.read()
            else:
                logger.warning(f"テンプレートファイルが見つかりません: {template_path}")
                return None
        except Exception as e:
            logger.error(f"テンプレート読み込みエラー: {e}")
            return None
    
    def _is_internal_email(self, from_address):
        """社内メールかどうかを判定"""
        if not from_address or not Config.INTERNAL_DOMAINS:
            return False
        
        from_address_lower = from_address.lower()
        for domain in Config.INTERNAL_DOMAINS:
            if domain.strip() and domain.strip().lower() in from_address_lower:
                return True
        return False
    
    def _extract_sender_info(self, from_address):
        """送信者情報を抽出（名前と会社名）"""
        # 形式: "山田太郎 <yamada@example.com>" または "yamada@example.com"
        sender_name = ""
        sender_company = ""
        
        if '<' in from_address:
            # 名前部分を抽出
            sender_name = from_address.split('<')[0].strip()
        
        # 会社名はメールアドレスから推測（簡易版）
        # 実際の会社名は別途取得が必要
        
        return sender_name, sender_company
    
    def _get_default_template(self, is_internal, sender_name, sender_company, my_name):
        """デフォルトのテンプレートを返す（テンプレートファイルが読めない場合のフォールバック）"""
        if is_internal:
            return f"""
【返信文のフォーマット】
以下の形式で返信文を作成してください：

{sender_name if sender_name else '〇〇'}様

お疲れさまです、{my_name}です。

（本文）

よろしくお願いいたします。
"""
        else:
            company_name = Config.MY_COMPANY_NAME if Config.MY_COMPANY_NAME else "株式会社〇〇"
            return f"""
【返信文のフォーマット】
以下の形式で返信文を作成してください：

{sender_company if sender_company else '〇〇株式会社'}　{sender_name if sender_name else '〇〇'}様

お世話になっております。
{company_name}の{my_name}です。

（本文）

以上、よろしくお願いいたします。
"""
    
    def generate_reply(self, original_subject, original_body, from_address=""):
        """返信文を生成（文体学習機能付き、社内外テンプレート対応）"""
        logger.info(f"返信文生成開始: {original_subject[:50]}...")
        
        # 社内/社外判定
        is_internal = self._is_internal_email(from_address)
        sender_name, sender_company = self._extract_sender_info(from_address)
        
        logger.info(f"メール種別: {'社内' if is_internal else '社外'}")
        
        # テンプレートを読み込んで変数を埋め込む
        if is_internal:
            # 社内用テンプレート
            template_path = Config.INTERNAL_EMAIL_TEMPLATE
            my_name = Config.MY_NAME_INTERNAL if Config.MY_NAME_INTERNAL else Config.MY_NAME
        else:
            # 社外用テンプレート
            template_path = Config.EXTERNAL_EMAIL_TEMPLATE
            my_name = Config.MY_NAME if Config.MY_NAME else "〇〇"
        
        # テンプレートファイルを読み込む
        template = self._load_template(template_path)
        
        if template:
            # テンプレート内の変数を置き換えてフォーマット指示を作成
            template_vars = {
                'sender_name': sender_name if sender_name else '〇〇',
                'sender_company': sender_company if sender_company else '〇〇株式会社',
                'my_name': my_name,
                'my_company': Config.MY_COMPANY_NAME if Config.MY_COMPANY_NAME else "株式会社〇〇",
                'body': '（本文）'
            }
            
            try:
                template_preview = template.format(**template_vars)
                template_instruction = f"""
【返信文のフォーマット】
以下の形式で返信文を作成してください：

{template_preview}

注意: （本文）の部分に、元のメールに対する適切な返信内容を記述してください。
"""
            except KeyError as e:
                logger.warning(f"テンプレート変数エラー: {e}")
                # フォールバック: デフォルトテンプレート
                template_instruction = self._get_default_template(is_internal, sender_name, sender_company, my_name)
        else:
            # テンプレートが読み込めない場合はデフォルトを使用
            logger.info("デフォルトテンプレートを使用します")
            template_instruction = self._get_default_template(is_internal, sender_name, sender_company, my_name)
        
        # プロンプトの基本部分
        prompt = f"""
以下のメールに対する適切な返信文を作成してください。

{template_instruction}
"""
        
        # 文体サンプルがある場合は、プロンプトに追加
        if self.writing_style_samples and len(self.writing_style_samples) > 0:
            logger.debug(f"文体サンプル {len(self.writing_style_samples)}件を使用します")
            prompt += f"""
【重要】以下は私が過去に送信したメールの文体サンプルです。
この文体・トーン・表現方法を参考にして、同じような文体で返信文を作成してください。

=== 文体サンプル ===
"""
            # 最大3件のサンプルを含める
            for idx, sample in enumerate(self.writing_style_samples[:3], 1):
                prompt += f"\n【サンプル{idx}】\n{sample}\n"
            
            prompt += "\n=== サンプル終了 ===\n\n"
        else:
            prompt += "丁寧で簡潔な返信を心がけてください。\n\n"
        
        prompt += f"""
【返信対象メール】
件名: {original_subject}

本文:
{original_body[:1000]}

上記のメールに対する返信文を、{('上記の文体サンプルを参考にして、' if self.writing_style_samples else '')}作成してください:
"""
        
        try:
            system_message = "あなたはビジネスメールの専門家です。"
            if self.writing_style_samples:
                system_message += "ユーザーの文体を忠実に再現してください。"
            
            response = self.client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": system_message},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=1000
            )
            logger.info(f"返信文生成完了: {original_subject[:50]}...")
            return response.choices[0].message.content
        except Exception as e:
            logger.error(f"返信文生成エラー ({original_subject[:50]}...): {e}", exc_info=True)
            return None
    
    def check_abnormal(self, subject, body):
        """異常を検知"""
        text = f"{subject} {body}"
        text_lower = text.lower()
        
        # 正常キーワードチェック
        for keyword in Config.NORMAL_KEYWORDS:
            if keyword.lower() in text_lower:
                return 'normal'
        
        # 異常キーワードチェック
        for keyword in Config.ABNORMAL_KEYWORDS:
            if keyword.lower() in text_lower:
                return 'abnormal'
        
        return 'unknown'
    
    def extract_todos(self, body):
        """TODOを抽出"""
        logger.info("TODO抽出開始")
        prompt = f"""
以下のメール内容から、人間が行う必要があるタスクやTODOを抽出してください。
- [ ] 形式でリストアップしてください。

本文:
{body[:1000]}
"""
        try:
            response = self.client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "あなたはタスク管理の専門家です。"},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=500
            )
            logger.info("TODO抽出完了")
            return response.choices[0].message.content
        except Exception as e:
            logger.error(f"TODO抽出エラー: {e}", exc_info=True)
            return None

