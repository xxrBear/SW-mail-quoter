import xlwings as xw
from bs4 import BeautifulSoup

from core.schemas import EachMail
from processor.base import ProcessorStrategy
from processor.mapping import (
    CBG_EXCEL_PROCESSING_RULES_MAPPING,
    CBG_QUOTE_FIELD_MAPPING,
)


class CustomerCBGProcessor(ProcessorStrategy):
    def process_excel(self, mail: EachMail, wb: xw.Book) -> float:
        """
        操作 Excel 文件，获取指定表格中的值
        :param mail: EachMail 对象
        :param wb: xlwings 工作簿对象
        :return: quote_value: 从 Excel 中获取的经过处理的报价值
        """

        try:
            sheet = wb.sheets[mail.sheet_name]

            target, excel_rules_mapping = self.get_excel_processing_rules(
                mail.sheet_name
            )
            # 将指定邮件内容写入 Excel

            for header, value in mail.df_dict.items():
                if excel_rules_mapping.get(header):
                    cell, transform = excel_rules_mapping[header]
                    sheet.range(cell).value = transform(value)

            quote_value = sheet.range(target).value

            wb.save()
        except Exception as e:
            print("操作 Excel 失败：", e)
            wb.close()

        return quote_value

    def process_mail_html(self, mail: EachMail, quote_value: float):
        """
        处理邮件 HTML 内容
        :param mail: EachMail 对象
        :param quote_value: 从 Excel 中获取的报价值
        :return: 修改后的 mail
        """
        quoted_field = self.get_quoted_field(mail.sheet_name)

        for label, td in self.iter_label_rows(mail.soup):
            if label == quoted_field:
                p = td.select_one("p")
                if p:
                    p.string = f"{quote_value:.2%}"
                    print(f"已修改报价字段 {label} 为：{quote_value:.2%}")
                break
        mail.content.html = str(mail.soup)
        return mail

    def get_excel_processing_rules(self, sheet_name: str) -> dict:
        """
        获取需要处理的 Excel 中表格的对应位置和方法
        :param sheet_name: 工作表名称
        :return: 映射字典
        """
        excel_processing_rules = CBG_EXCEL_PROCESSING_RULES_MAPPING.get(sheet_name)
        if not excel_processing_rules:
            raise ValueError(f"未找到对应的 Excel 规则映射，工作表: {sheet_name}")
        return excel_processing_rules

    def is_already_quoted(self, df_dict: dict, sheet_name: str) -> bool:
        """
        判断邮件中是否已完成报价

        已完成报价的逻辑：
        - 所有字段值都非空
        - 或需报价字段与系统记录不一致

        :param df_dict: 邮件中提取的表格数据（字典格式）
        :param sheet_name: 表格名称，用于获取需报价字段
        :return: True 表示已报价，False 表示未报价
        """
        # 判断是否所有值都非空
        all_fields_filled = all(df_dict.values())

        # 获取系统记录的需报价字段（字符串）
        required_fields = self.get_quoted_field(sheet_name)

        # 提取实际为空的字段名组成字符串
        actual_empty_fields = "".join(
            str(k).strip() for k, v in df_dict.items() if v is None
        )

        # 判断是否与系统记录一致
        field_mismatch = required_fields != actual_empty_fields

        return all_fields_filled or field_mismatch

    def get_quoted_field(self, sheet_name: str) -> str:
        """
        获取需要报价的字段
        :param sheet_name: 工作表名称
        :return: 需要报价的字段名称
        """
        return CBG_QUOTE_FIELD_MAPPING.get(sheet_name)

    def iter_label_rows(self, soup: BeautifulSoup):
        """返回需要处理的标签行"""
        for row in soup.select("table tr"):
            tds = row.find_all("td", recursive=False)
            if len(tds) == 2:
                label = tds[0].get_text(strip=True)
                yield label, tds[1]
