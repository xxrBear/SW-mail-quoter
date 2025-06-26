from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from typing import Dict, List

import xlwings as xw

from core.client import mail_client, send_mail_client
from core.context import mail_context
from core.excel import ExcelHandler
from core.schemas import EachMail
from core.utils import print_banner
from db.models import MailState
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

        # 过滤不可报价结果字典
        filter_dict = self.filter_unquotable_result_dict(result_dict)

        # 处理未报价邮件并回复
        mail_state = MailState()
        excel_handler = ExcelHandler()

        pending_emails = []
        for eamil_addr, result_list in filter_dict.items():
            print_banner("开始处理可报价邮件......")
            processor = get_processor(eamil_addr)  # 获取每个客户对应的邮件处理策略

            # 有可报价邮件，清空 Sheet 行
            sheet_set = set([i.sheet_name for i in result_list if i.sheet_name])
            for _sheet_name in sheet_set:
                excel_handler.clear_sheet_columns(wb, _sheet_name)

            sheet_name_count_dict = {_sheet_name: 0 for _sheet_name in sheet_set}

            for mail in result_list:
                print(f"处理邮件: {mail.subject} 来自: {eamil_addr}")

                # 处理 Excel 对应 Sheet
                excel_handler.copy_sheet_columns(
                    wb, mail.sheet_name, sheet_name_count_dict[mail.sheet_name]
                )

                # 获取报价值，并写入待发送邮件内容中
                quote_value = processor.process_excel(
                    mail, wb, sheet_name_count_dict[mail.sheet_name]
                )
                processed_mail = processor.process_mail_html(mail, quote_value)

                if not (
                    processed_mail.underlying.startswith("AU")
                    or processed_mail.underlying.startswith("XAU")
                ):
                    print(
                        f"非 AU 开头标的合约 {processed_mail.underlying}，暂时跳过 {processed_mail.subject} \n"
                    )
                    excel_handler.delete_sheet_column(
                        wb, mail.sheet_name, sheet_name_count_dict[mail.sheet_name]
                    )
                    self.skip(mail, "非 AU 或 XAU 开头的标的合约，暂时跳过")
                    continue

                # 待回复邮件内容
                pending_emails.append(processed_mail)

                # 写入数据库
                try:
                    mail_state.update_or_create_record(processed_mail)  # type: ignore
                except Exception as e:
                    print(f"写入数据库出错: {e}")

                sheet_name_count_dict[mail.sheet_name] += 1

        # 使用多线程发送邮件
        # with ThreadPoolExecutor(max_workers=10) as executor:
        #     futures = [
        #         executor.submit(send_mail_client.reply_mail, p) for p in pending_emails
        #     ]
        #     for f in as_completed(futures):
        #         try:
        #             f.result()
        #         except Exception as e:
        #             print(f"发送失败: {e}")

        # 处理异常邮件，写入 Excel
        try:
            excel_handler.process_abnormal_mails_sheet(wb)
        except Exception as e:
            print(f"写入今日报价异常报错: {e}")

        # 今日成功报价数据写入 Excel
        try:
            excel_handler.process_successful_mails_sheet(wb)
        except Exception as e:
            print(f"写入今日成功报价报错：{e}")

        # 处理 hold价
        try:
            excel_handler.process_hold_mails_sheet(wb)
        except Exception as e:
            print(f"写入hold价失败：{e}")

    def filter_unquotable_result_dict(
        self, result_dict: Dict[str, List[EachMail]]
    ) -> Dict[str, List[EachMail]]:
        """过滤不可报价的邮件，并由上下文对象记录"""
        filtered_dict = defaultdict(list)

        mail_state = MailState()  # 数据库表

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

                if mail_state.mail_exists(mail):
                    continue

                filtered_dict[email_addr].append(mail)

        return filtered_dict

    def skip(self, mail: EachMail, reason: str):
        mail_context.skip_mail(
            mail.subject, mail.from_addr, mail.sent_time, datetime.now(), reason
        )
