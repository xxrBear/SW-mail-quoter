import pandas as pd
import xlwings as xw
from bs4 import BeautifulSoup

from models.schemas import EachMail
from processor.base import ProcessorStrategy
from processor.mapping import CBG_BULL_LADDER_TUPLE, CBG_BINARRY_CALL_TUPLE


class CustomerCBGProcessor(ProcessorStrategy):
    def process_excel(self, df: pd.DataFrame, wb: xw.Book, sheet_name: str) -> float:
        """
        操作 Excel 文件
        :param df: 解析后的 DataFrame
        :return: k1: 从 Excel 中获取的值
        """

        try:
            sheet = wb.sheets[sheet_name]

            target, excel_mapping = self.get_except_excel_mapping(sheet_name)
            # 将指定邮件内容写入 Excel
            for _, column in df.iterrows():
                header, value = column
                if header in excel_mapping:
                    cell, transform = excel_mapping[header]
                    sheet.range(cell).value = transform(value)

            k1 = sheet.range(target).value

            wb.save()
        except Exception as e:
            print("操作 Excel 失败：", e)
            wb.close()

        return k1

    def process_mail_html(self, mail: EachMail, k1: float):
        """
        处理邮件 HTML 内容
        :param mail: 邮件对象
        :param k1: 从 Excel 中获取的值
        :return: 修改后的 mail
        """

        soup = BeautifulSoup(mail.content.html, "html.parser")
        # 查找所有表格行
        for row in soup.select("table tr"):
            tds = row.find_all("td", recursive=False)
            if len(tds) != 2:
                continue

            label = tds[0].get_text(strip=True)
            if label in ("行权价格1（低）", "行权价格"):
                span = tds[1].select_one("p > span")
                if span:
                    span.string = f"{k1:.2%}"
                    print(f"已修改 {label} 为：{k1:.2%}")
                break

        mail.content.html = str(soup)
        return mail

    def get_except_excel_mapping(self, sheet_name: str) -> dict:
        """
        获取需要处理的 Excel 中表格的对应位置和方法
        :param sheet_name: 工作表名称
        :return: 映射字典
        """
        if sheet_name == "看涨阶梯":
            return CBG_BULL_LADDER_TUPLE
        elif sheet_name == "二元看涨":
            return CBG_BINARRY_CALL_TUPLE
        else:
            raise ValueError(f"未找到对应的工作表对应操作，工作表: {sheet_name}")
