import xlwings as xw
from bs4 import BeautifulSoup

from core.schemas import EachMail
from core.utils import (
    add_excel_subject_cell,
    calc_next_letter,
    get_rate,
    get_risk_free_rate,
)
from processor.base import ProcessorStrategy
from processor.mapping import get_sheet_handler


class CustomerCBGProcessor(ProcessorStrategy):
    def process_excel(
        self, mail: EachMail, wb: xw.Book, sheet_copy_count: int
    ) -> float:
        """
        操作 Excel 文件，获取指定表格中的值
        :param mail: EachMail 对象
        :param wb: xlwings 工作簿对象
        :return: quote_value: 从 Excel 中获取的经过处理的报价值
        """

        try:
            sheet = wb.sheets[mail.sheet_name]

            # 对应邮件中的数据
            sheet_mapping_handler = get_sheet_handler(mail.sheet_name)
            # Excel 待处理字段
            fields_to_update = sheet_mapping_handler.fields_rule_dict

            # 将指定邮件内容写入 Excel
            next_letter = calc_next_letter("C", sheet_copy_count)
            for header, value in mail.df_dict.items():
                if fields_to_update.get(header):
                    cell, apply_method = fields_to_update[header]
                    finally_cell = next_letter + str(cell)
                    sheet.range(finally_cell).value = apply_method(value)

            # 处理 Excel 中的数据
            other_dict = sheet_mapping_handler.other_dict

            # 交易日
            trade_date_index = sheet.range(next_letter + other_dict.get("交易日"))
            trade_date_formula = trade_date_index.formula
            if str(mail.underlying).startswith("AU"):
                trade_date_formula = trade_date_formula.replace("$C", "$A")
                trade_date_index.formula = trade_date_formula

            T_ = sheet.range(next_letter + other_dict.get("T")).value

            # VOL
            rate = get_rate(mail.underlying, T_, wb)
            sheet.range(next_letter + other_dict.get("VOL")).value = rate

            # 无风险利率
            r = get_risk_free_rate(mail.underlying)
            sheet.range(next_letter + other_dict.get("无风险利率")).value = r

            # 获取需报价字段所在位置并读取
            finally_target = next_letter + str(sheet_mapping_handler.quote_line)
            quote_value = sheet.range(finally_target).value

            # 每个表格底部添加邮件标题和哈希值
            add_excel_subject_cell(wb, mail, next_letter)

        except Exception as e:
            print("操作 Excel 失败：", e)

        return quote_value

    def process_mail_html(self, mail: EachMail, quote_value: float):
        """
        处理邮件 HTML 内容
        :param mail: EachMail 对象
        :param quote_value: 从 Excel 中获取的报价值
        :return: 修改后的 mail
        """
        sheet_mapping_handler = get_sheet_handler(mail.sheet_name)
        quote_name = sheet_mapping_handler.quote_name

        for label, td in self.iter_label_rows(mail.soup):
            if label == quote_name:
                p = td.select_one("p")
                if p:
                    quote_value = f"*{quote_value:.3f}" if quote_value else "*0.000"
                    p.string = quote_value
                    print(f"已修改报价字段 {label} 为：{quote_value} \n")
                break
        mail.content.html = str(mail.soup)
        return mail

    def cannot_quote(self, mail: EachMail) -> bool:
        """
        判断邮件中是否不满足报价条件

        不能报价的邮件所需条件：
        - 所有字段值都非空
        - 或需报价字段与系统配置的不一致

        :param df_dict: 邮件中提取的表格数据（字典格式）
        :param sheet_name: 表格名称，用于获取需报价字段
        :return: True 表示满足，False 表示不满足
        """
        # 判断是否所有值都非空
        all_fields_filled = all(mail.df_dict.values())

        # 获取系统记录的需报价字段（字符串）
        quote_name = get_sheet_handler(mail.sheet_name).quote_name

        # 提取实际为空的字段名组成字符串
        actual_empty_fields = "".join(
            str(k).strip() for k, v in mail.df_dict.items() if v is None
        )

        # 判断是否与系统记录一致
        field_mismatch = quote_name != actual_empty_fields

        return all_fields_filled or field_mismatch

    def iter_label_rows(self, soup: BeautifulSoup):
        """返回需要处理的标签行"""
        if isinstance(soup, str):
            soup = BeautifulSoup(soup, "html.parser")

        for row in soup.select("table tr"):
            tds = row.find_all("td", recursive=False)
            if len(tds) == 2:
                label = tds[0].get_text(strip=True)
                yield label, tds[1]
