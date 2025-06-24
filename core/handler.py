from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from typing import Dict, List

import xlwings as xw

from core.client import mail_client, send_mail_client
from core.context import mail_context
from core.schemas import EachMail
from core.utils import calc_next_letter, print_banner
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

                if not processed_mail.underlying.startswith("AU"):
                    print("非 AU 开头标的合约，暂时跳过")
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
        #             print(
        #                 f"已回复邮件: {processed_mail.subject} 来自: {eamil_addr} \n "
        #             )
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


class ExcelHandler:
    """
    Excel 公共规则处理类
    """

    def clear_sheet_columns(self, wb: xw.Book, sheet_name: str) -> None:
        """首次处理时，清空对应表格的列"""
        sheet = wb.sheets[sheet_name]
        sheet.range("C:Z").delete()  # 清除值、格式、批注等
        wb.save()

    def copy_sheet_columns(
        self, wb: xw.Book, sheet_name: str, sheet_copy_count: int
    ) -> None:
        """复制工作表的列"""
        sheet = wb.sheets[sheet_name]
        letter = calc_next_letter("C", sheet_copy_count)

        # --- 禁用 Excel 事件和显示警告 ---
        wb.app.enable_events = False  # 禁用VBA事件
        wb.app.display_alerts = False  # 禁用Excel自身的警告弹窗
        try:
            # 从 B 列复制一百行
            sheet.range("B1:B100").api.Copy(
                Destination=sheet.range(f"{letter}1:{letter}100").api
            )
            sheet.api.Application.CutCopyMode = False  # 清除 复制模式 的虚线框

        finally:
            # 无论代码是否出错，都确保这些设置被恢复，否则会影响后续的Excel操作
            wb.app.enable_events = True
            wb.app.display_alerts = True
            sheet.api.Application.CutCopyMode = True
            wb.save()

    def ensure_sheet_exists(self, wb: xw.Book, sheet_name: str):
        """确保 sheet_name Sheet 存在，如果不存在就创建"""
        if sheet_name not in [s.name for s in wb.sheets]:  # type: ignore
            sheet = wb.sheets.add(name=sheet_name, before="8080结构")  # type: ignore
            self._init_sheet_header(sheet)
        else:
            sheet = wb.sheets[sheet_name]  # type: ignore
        return sheet

    def _init_sheet_header(self, sheet: xw.Sheet):
        """设置Sheet 表头和样式"""
        if "失败" in sheet.name:
            sheet.range("A1").value = "邮件主题"
            sheet.range("B1").value = "失败原因"
            sheet.range("C1").value = "发件人"
            sheet.range("D1").value = "询价时间"
            sheet.range("E1").value = "报价时间"
            # 表格颜色
            sheet.api.Tab.Color = 255  # 红色
        else:
            sheet.range("A1").value = "邮件主题"
            sheet.range("B1").value = "发件人"
            sheet.range("C1").value = "询价时间"
            sheet.range("D1").value = "报价时间"
            # 表格颜色
            sheet.api.Tab.Color = 65280  # 绿色

        # 定位表头范围
        header_range = sheet.range("A1").expand("right")
        # 设置样式
        header_range.api.Font.Bold = True  # 加粗
        header_range.api.HorizontalAlignment = -4108  # 水平居中
        header_range.api.VerticalAlignment = -4108  # 垂直居中
        header_range.color = (192, 192, 192)  # 设置背景色
        header_range.columns.autofit()
        header_range.column_width = 80  # 表头宽度
        header_range.row_height = 30  # 表头高度

    def clear_sheet_content(self, sheet: xw.Sheet):
        """清空指定 Sheet 表头以下所有内容"""
        used_range = sheet.used_range
        if used_range.rows.count > 1:
            # 清空 A2:最后一行最后一列
            last_cell = used_range.end("down").end("right")
            sheet.range(f"A2:{last_cell.address}").clear()

    def write_abnormal_mails(self, sheet: xw.Sheet):
        """把上下文中的异常邮件批量写入 Sheet"""
        print_banner("开始写入报价失败的邮件数据...")

        if not mail_context.email:
            return

        data = [
            [
                mail.get("subject"),
                mail.get("reason"),
                mail.get("sent_addr"),
                mail.get("sent_time"),
                mail.get("created_time"),
            ]
            for mail in mail_context.email
        ]
        # 从 A2 开始批量写值
        sheet.range("A2").value = data

    def write_today_successful_mails(self, sheet: xw.Sheet):
        print_banner("开始写入报价成功的邮件数据...")

        mail_state = MailState()
        result = mail_state.get_successful_mail_info()  # type: ignore
        sheet.range("A2").value = result

    def process_abnormal_mails_sheet(self, wb: xw.Book):
        """处理异常邮件的 sheet"""
        sheet = self.ensure_sheet_exists(wb, "今日失败报价")
        self.clear_sheet_content(sheet)
        self.write_abnormal_mails(sheet)
        wb.save()

    def process_successful_mails_sheet(self, wb: xw.Book):
        sheet = self.ensure_sheet_exists(wb, "今日成功报价")
        self.clear_sheet_content(sheet)
        self.write_today_successful_mails(sheet)
        wb.save()
