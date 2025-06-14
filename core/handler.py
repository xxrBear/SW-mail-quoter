from collections import defaultdict
from datetime import date
from typing import Dict, List

import xlwings as xw

from core.client import mail_client
from core.schemas import EachMail
from core.utils import print_banner
from db.models import MailState, MailStateEnum
from processor.registry import get_processor


class MailHandler:
    def __init__(self, folder: str = "INBOX", since_date: date = date.today()) -> None:
        self.folder = folder
        self.since_date = since_date

    def handle(self, wb: xw.Book) -> None:
        # 读取邮件并获取结果字典
        result_dict = mail_client.read_mail(
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

                # 第一次处理时清除 Excel 中的值
                sheet_name_count = MailState().count_sheet_name(mail)
                if not sheet_name_count:
                    self.clear_sheet_columns(wb, mail.sheet_name)
                else:
                    self.copy_sheet_columns(wb, mail.sheet_name, sheet_name_count)

                quote_value = processor.process_excel(mail, wb, sheet_name_count)
                processed_mail = processor.process_mail_html(mail, quote_value)

                # 回复邮件
                mail_client.reply_mail(processed_mail)
                # 写入数据库
                MailState().update_mail_state(processed_mail, MailStateEnum.PROCESSED)
                print(f"已回复邮件: {processed_mail.subject} 来自: {eamil_addr}")
                print()

    def filter_quoted_result_dict(
        self, result_dict: Dict[str, List[EachMail]]
    ) -> Dict[str, List[EachMail]]:
        """排除已报价的邮件，并返回需要处理的邮件字典"""
        modify_dict = defaultdict(list)

        for email_addr, mails in result_dict.items():
            processor = get_processor(email_addr)
            if not processor:
                print(f"未找到对应的邮箱处理策略，邮箱地址: {email_addr}")
                continue

            for mail in mails:
                # 过滤已处理过的邮件
                if MailState().is_mail_exists(mail):
                    continue

                # 处理已报价邮件
                if processor.is_already_quoted(mail.df_dict, mail.sheet_name):
                    print(f"当前邮件已完成报价，跳过邮件: {mail.subject}")
                    MailState().update_mail_state(mail, MailStateEnum.MANUAL)
                    continue

                modify_dict[email_addr].append(mail)

        return modify_dict

    def clear_sheet_columns(self, wb: xw.Book, sheet_name: str) -> None:
        sheet = wb.sheets[sheet_name]
        sheet.range("C:Z").delete()  # 清除值、格式、批注等

    def copy_sheet_columns(
        self, wb: xw.Book, sheet_name: str, sheet_name_count: int
    ) -> None:
        sheet = wb.sheets[sheet_name]

        sheet.range("B:B").api.Copy(Destination=sheet.range("C:C").api)
