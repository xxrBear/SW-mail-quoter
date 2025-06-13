import hashlib
from collections import defaultdict
from datetime import date

import xlwings as xw

from core.client import create_mail_client
from core.schemas import EachMail
from db.engine import SessionLocal
from db.models import MailState
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

        # 处理结果字典，排除已报价的邮件
        filter_dict = self.filter_quoted_result_dict(result_dict)

        # 处理未报价邮件并回复
        print("开始处理未报价邮件......")
        print(
            "-------------------------------------------------------------------------------"
        )

        for eamil_addr, result_list in filter_dict.items():
            processor = get_processor(eamil_addr)  # 获取每个客户对应的邮件处理策略

            for mail in result_list:
                print(f"处理邮件: {mail.subject} 来自: {eamil_addr}")

                quote_value = processor.process_excel(mail, wb)
                processed_mail = processor.process_mail_html(mail, quote_value)

                # 回复邮件
                self.mail_client.reply_mail(processed_mail)
                # 写入数据库
                self.update_mail_state(processed_mail)
                print(f"已回复邮件: {processed_mail.subject} 来自: {eamil_addr}")

    def filter_quoted_result_dict(self, result_dict: dict) -> dict:
        """排除已报价的邮件，并返回需要处理的邮件字典
        :param result_dict: 原始邮件字典，键为发件人地址，值为 EachMail 对象列表
        :return: 返回一个新的字典，键为发件人地址，值为需要处理的 EachMail 对象列表
        """
        modify_dict = defaultdict(list)

        for email_addr, mails in result_dict.items():
            processor = get_processor(email_addr)
            if not processor:
                print(f"未找到对应的邮箱处理策略，邮箱地址: {email_addr}")
                continue

            for mail in mails:
                # 过滤已处理过的邮件
                if self.is_processed(mail):
                    continue

                # 处理已报价邮件
                if processor.is_already_quoted(mail.df_dict, mail.sheet_name):
                    print(f"当前邮件已完成报价，跳过邮件: {mail.subject}")
                    self.write_to_database(mail)
                    continue

                modify_dict[email_addr].append(mail)

        return modify_dict

    def write_to_database(self, mail: EachMail) -> None:
        """将处理结果写入数据库"""
        subject = mail.subject
        sent_time = mail.sent_time

        join_str = f"{subject} - {sent_time.strftime('%Y-%m-%d %H:%M:%S')}"
        hash_obj = hashlib.sha256(join_str.encode("utf-8"))
        mail_hash = hash_obj.hexdigest()

        mail_state = MailState(
            mail_hash=mail_hash,
        )

        with SessionLocal() as session:
            session.add(mail_state)
            session.commit()
            print(f"邮件状态已保存: {mail_hash}")

    def is_processed(self, mail: EachMail) -> bool:
        """检查邮件是否已处理"""
        join_str = f"{mail.subject} - {mail.sent_time.strftime('%Y-%m-%d %H:%M:%S')}"
        hash_obj = hashlib.sha256(join_str.encode("utf-8"))
        mail_hash = hash_obj.hexdigest()

        with SessionLocal() as session:
            return (
                session.query(MailState).filter_by(mail_hash=mail_hash).first()
                is not None
            )

    def update_mail_state(self, mail: EachMail) -> None:
        """更新邮件状态到数据库"""
        join_str = f"{mail.subject} - {mail.sent_time.strftime('%Y-%m-%d %H:%M:%S')}"
        hash_obj = hashlib.sha256(join_str.encode("utf-8"))
        mail_hash = hash_obj.hexdigest()

        mail_state = MailState(mail_hash=mail_hash)

        with SessionLocal() as session:
            session.add(mail_state)
            session.commit()
            print(f"邮件状态已更新: {mail_hash}")
