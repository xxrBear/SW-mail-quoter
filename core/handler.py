from collections import defaultdict
from datetime import date

import xlwings as xw

from core.client import create_mail_client
from core.parser import get_mail_hash
from core.schemas import EachMail
from core.utils import print_banner
from db.decorator import with_session
from db.engine import SessionLocal
from db.models import MailState, MailStateEnum
from processor.registry import get_processor


class MailHandler:
    def __init__(self, folder: str = "INBOX", since_date: date = date.today()) -> None:
        self.folder = folder
        self.since_date = since_date
        self.mail_client = create_mail_client()

    def handle(self, wb: xw.Book) -> dict:
        # 读取邮件并获取结果字典
        result_dict = self.mail_client.read_mail(
            folder=self.folder, since_date=self.since_date
        )

        # 排除已报价的邮件
        filter_dict = self.filter_quoted_result_dict(result_dict)

        # 处理未报价邮件并回复
        print_banner("开始处理未报价邮件......")
        for eamil_addr, result_list in filter_dict.items():
            processor = get_processor(eamil_addr)  # 获取每个客户对应的邮件处理策略

            for mail in result_list:
                print(f"处理邮件: {mail.subject} 来自: {eamil_addr}")

                quote_value = processor.process_excel(mail, wb)
                processed_mail = processor.process_mail_html(mail, quote_value)

                # 回复邮件
                self.mail_client.reply_mail(processed_mail)
                # 写入数据库
                self.update_mail_state(processed_mail, MailStateEnum.PROCESSED)
                print(f"已回复邮件: {processed_mail.subject} 来自: {eamil_addr}")
                print()

    def filter_quoted_result_dict(self, result_dict: dict) -> dict:
        """排除已报价的邮件，并返回需要处理的邮件字典"""
        modify_dict = defaultdict(list)

        for email_addr, mails in result_dict.items():
            processor = get_processor(email_addr)
            if not processor:
                print(f"未找到对应的邮箱处理策略，邮箱地址: {email_addr}")
                continue

            for mail in mails:
                # 过滤已处理过的邮件
                if self.is_mail_processed(mail):
                    continue

                # 处理已报价邮件
                if processor.is_already_quoted(mail.df_dict, mail.sheet_name):
                    print(f"当前邮件已完成报价，跳过邮件: {mail.subject}")
                    self.update_mail_state(mail, MailStateEnum.MANUAL)
                    continue

                modify_dict[email_addr].append(mail)

        return modify_dict

    @with_session
    def update_mail_state(
        session: SessionLocal, self, mail: EachMail, state: MailStateEnum
    ) -> None:
        """将处理结果写入数据库"""
        mail_hash = get_mail_hash(mail)
        mail_state = MailState(
            mail_hash=mail_hash,
        )

        session.add(mail_state)
        session.commit()

    @with_session
    def is_mail_processed(session: SessionLocal, self, mail: EachMail) -> bool:
        """检查邮件是否已处理"""
        mail_hash = get_mail_hash(mail)
        return (
            session.query(MailState).filter_by(mail_hash=mail_hash).first() is not None
        )
