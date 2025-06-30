from datetime import datetime

import xlwings as xw

from core.context import mail_context
from core.utils import (
    calc_next_letter,
    col_index_to_letter,
    find_position_in_column,
    print_banner,
)
from db.models import MailState


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
        elif "成功" in sheet.name:
            sheet.range("A1").value = "邮件主题"
            sheet.range("B1").value = "发件人"
            sheet.range("C1").value = "询价时间"
            sheet.range("D1").value = "报价时间"
            # 表格颜色
            sheet.api.Tab.Color = 65280  # 绿色
        else:
            sheet.range("A1").value = "邮件主题"
            sheet.range("B1").value = "发件人"
            sheet.range("C1").value = "询价时间"
            sheet.range("D1").value = "报价时间"
            sheet.api.Tab.Color = 65535  # 黄色

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
        print_banner("开始写入当次报价失败的邮件数据...")

        if not mail_context.email:
            return

        data = [
            [
                mail.get("subject"),
                mail.get("reason"),
                mail.get("sent_addr"),
                mail.get("sent_time"),
                datetime.now(),
            ]
            for mail in mail_context.email
        ]
        # 从 A2 开始批量写值
        sheet.range("A2").value = data

    def write_today_successful_mails(self, sheet: xw.Sheet):
        print_banner("开始写入今日报价成功的邮件数据...")

        mail_state = MailState()
        result = mail_state.get_successful_mail_info()  # type: ignore
        sheet.range("A2").value = result

    def write_hold_mails(self, sheet: xw.Sheet):
        print_banner("开始写入当次hold价的邮件数据...")

        if not mail_context.hold_email:
            return

        data = [
            [
                mail.get("subject"),
                mail.get("sent_addr"),
                mail.get("sent_time"),
                datetime.now(),
            ]
            for mail in mail_context.hold_email
        ]
        # 从 A2 开始批量写值
        sheet.range("A2").value = data

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

    def process_hold_mails_sheet(self, wb: xw.Book):
        sheet = self.ensure_sheet_exists(wb, "hold价邮件")
        self.clear_sheet_content(sheet)
        self.write_hold_mails(sheet)
        wb.save()

    @classmethod
    def get_confirmed_mail_subject(cls, sheet: xw.Sheet):
        value = "是否可以回复报价邮件（是/否）"
        row, _ = find_position_in_column(sheet, value, "A")
        if not row:
            return

        cell_range = sheet.range(f"C{row}:Z{row}")
        subjects = []
        for cell in cell_range:
            if str(cell.value).strip() == "是":
                col = col_index_to_letter(cell.column)
                target = f"{col}{row + 1}"
                print(target)
                subjects.append(sheet.range(target).value)
        return subjects
