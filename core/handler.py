from collections import defaultdict
from datetime import date, datetime
from typing import Dict, List

import xlwings as xw

from core.client import mail_client
from core.context import mail_context
from core.excel import ExcelHandler
from core.schemas import EachMail
from core.utils import print_banner
from db.enums import MailStateEnum
from db.models import MailState
from processor.registry import get_processor, subject_sheet_map


class MailHandler:
    def __init__(self, folder: str = "INBOX", since_date: date = date.today()) -> None:
        self.folder = folder
        self.since_date = since_date

    def handle(self, wb: xw.Book) -> None:
        # 读取邮件并获取结果字典
        result_dict = mail_client.read_mail(
            folder=self.folder, since_date=self.since_date
        )

        # 过滤不可报价结果字典
        filter_dict = self.filter_unquotable_result_dict(result_dict)

        # 处理未报价邮件并回复
        excel_handler = ExcelHandler()

        # 清空Sheet
        sheet_names = subject_sheet_map.keys()
        for _sheet_name in sheet_names:
            excel_handler.clear_sheet_columns(wb, _sheet_name)

        for email_addr, result_list in filter_dict.items():
            print_banner("开始处理可报价邮件......")
            processor = get_processor(email_addr)  # 获取每个客户对应的邮件处理策略

            sheet_name_count_dict = {_sheet_name: 0 for _sheet_name in sheet_names}
            for mail in result_list:
                print(f"处理邮件: {mail.subject} 来自: 【{email_addr}】")

                # 处理 Excel 对应 Sheet
                excel_handler.copy_sheet_columns(
                    wb, mail.sheet_name, sheet_name_count_dict[mail.sheet_name]
                )

                # 获取报价值，并写入待发送邮件内容中
                processor.process_excel(
                    mail, wb, sheet_name_count_dict[mail.sheet_name]
                )

                # 写入数据库
                try:
                    MailState().create_record(mail)  # type: ignore
                except Exception as e:
                    print(f"写入数据库出错: {e}")

                sheet_name_count_dict[mail.sheet_name] += 1

        # 写入当次报价异常邮件
        try:
            excel_handler.process_abnormal_mails_sheet(wb)
        except Exception as e:
            print(f"写入今日报价异常报错: {e}")

        # 写入当次 hold价数据
        try:
            excel_handler.process_hold_mails_sheet(wb)
        except Exception as e:
            print(f"写入hold价失败：{e}")

        # 写入今日成功报价数据
        try:
            ExcelHandler().process_successful_mails_sheet(wb)
        except Exception as e:
            print(f"写入今日成功报价报错：{e}")

    def filter_unquotable_result_dict(
        self, result_dict: Dict[str, List[EachMail]]
    ) -> Dict[str, List[EachMail]]:
        """过滤不可报价的邮件，并由上下文对象记录"""
        filtered_dict = defaultdict(list)

        for email_addr, mails in result_dict.items():
            processor = get_processor(email_addr)

            for mail in mails:
                if not processor:
                    self.skip(mail, "未找到对应的邮箱处理策略")
                    continue

                # 处理不满足报价条件的邮件
                if processor.cannot_quote(mail):
                    self.skip(mail, "当前邮件不满足报价条件，跳过邮件")
                    continue

                db_mail = MailState().mail_exists(mail)
                if db_mail and db_mail.state != MailStateEnum.UNPROCESSED:
                    continue

                filtered_dict[email_addr].append(mail)

        return filtered_dict

    def skip(self, mail: EachMail, reason: str):
        mail_context.skip_mail(
            mail.subject, mail.from_addr, mail.sent_time, datetime.now(), reason
        )

    # ---------------------------------------------------------------------------------
    # CLI 指定方法
    # ---------------------------------------------------------------------------------

    def pull_quote_mails_to_db(self, since_date: date = date.today()):
        """获取报价邮件数据，存入数据库表中"""

        result_dict = mail_client.read_mail(folder=self.folder, since_date=since_date)
        filter_dict = self.filter_unquotable_result_dict(result_dict)

        for _, result_list in filter_dict.items():
            for each_mail in result_list:
                MailState().create_record(each_mail)
