import re

import pandas as pd
import xlwings as xw
from bs4 import BeautifulSoup

from models.schemas import EachMail
from processor.base import ProcessorStrategy


class CustomerCBGProcessor(ProcessorStrategy):
    excel_mapping = {
        "挂钩标的合约": (
            "C3",
            lambda v: re.findall(r"[（(](.*?)[）)]", v)[0].replace(".", "").upper(),
        ),
        "产品启动日": ("C4", str),
        "交割日（双方资金清算日）": ("C5", str),
        "最低收益率（年化）": ("C9", str),
        "中间收益率（年化）": ("C10", str),
        "最高收益率（年化）": ("C11", str),
        "行权价格2（高）": ("C22", lambda v: v.replace("*", "")),
    }

    def process_excel(self, df: pd.DataFrame, wb: xw.Book, sheet_name: str) -> float:
        """
        操作 Excel 文件
        :param df: 解析后的 DataFrame
        :return: k1: 从 Excel 中获取的值
        """

        try:
            sheet = wb.sheets[sheet_name]

            # 将指定邮件内容写入 Excel
            for _, column in df.iterrows():
                header, value = column
                if header in self.excel_mapping:
                    cell, transform = self.excel_mapping[header]
                    sheet.range(cell).value = transform(value)

            k1 = sheet.range("C23").value

            wb.save()
        except Exception as e:
            print("操作 Excel 失败：", e)
        finally:
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

        # 查找所有表格行 <tr>
        for row in soup.select("table tr"):
            tds = row.find_all("td", recursive=False)
            if len(tds) == 2 and tds[0].get_text(strip=True) == "行权价格1（低）":
                span = tds[1].select_one("p > span")
                if span:
                    span.string = "{:.2%}".format(k1)
                    print(f"已修改 行权价格1（低） 为：{k1}")
                break  # 找到后即可退出循环

        mail.content.html = str(soup.prettify())
        return mail
