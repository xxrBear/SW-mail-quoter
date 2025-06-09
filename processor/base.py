import re
from abc import ABC, abstractmethod

import pandas as pd
import xlwings as xw
from bs4 import BeautifulSoup

from models.schemas import EachMail


class ProcessorStrategy(ABC):
    @abstractmethod
    def process_excel(self):
        raise NotImplementedError()

    @abstractmethod
    def process_mail_html(self):
        raise NotImplementedError()


class CustomerAProcessor(ProcessorStrategy):
    def process_excel(self, df: pd.DataFrame) -> float:
        """
        操作 Excel 文件
        :param df: 解析后的 DataFrame
        :return: k1: 从 Excel 中获取的值
        """
        # 这里可以添加对 DataFrame 的处理逻辑

        # 启动 Excel 应用
        app = xw.App(visible=False, add_book=False)

        # 写入 Excel
        try:
            wb = app.books.open("./test.xlsm")
            sheet = wb.sheets["看涨阶梯"]

            # 将读出来的邮件内容写入 Excel
            for _, column in df.iterrows():
                header, value = column
                if header == "挂钩标的合约":
                    pattern = r"[（(](.*?)[）)]"
                    value2 = re.findall(pattern, value)
                    sheet.range("C3").value = value2[0].replace(".", "").upper()
                elif header == "产品启动日":
                    sheet.range("C4").value = value
                elif header == "交割日（双方资金清算日）":
                    sheet.range("C5").value = value
                elif header == "最低收益率（年化）":
                    sheet.range("C9").value = value
                elif header == "中间收益率（年化）":
                    sheet.range("C10").value = value
                elif header == "最高收益率（年化）":
                    sheet.range("C11").value = value
                elif header == "行权价格2（高）":
                    sheet.range("C22").value = value.replace("*", "")
                else:
                    pass

            k1 = sheet.range("C23").value

            wb.save()
        except Exception as e:
            print("操作 Excel 失败：", e)
        finally:
            wb.close()
            app.quit()

        return k1

    def process_mail_html(self, mail: EachMail, df: pd.DataFrame, k1: float):
        """
        处理邮件 HTML 内容
        :param mail_client: 邮箱客户端实例
        :param last_mail: 最后一封邮件实例
        :param k1: 从 Excel 中获取的值
        :return: None
        """

        html_content = mail.content.html
        soup = BeautifulSoup(html_content, "html.parser")

        # 查找所有表格行 <tr>
        rows = soup.select("table tr")

        # 遍历每一行，查找“行权价格1（低）”这一项并修改值
        for row in rows:
            tds = row.find_all("td", recursive=False)
            if len(tds) == 2:
                key = tds[0].get_text(strip=True)
                if key == "行权价格1（低）":
                    # 修改第二个单元格的内容
                    new_value = k1
                    tds[1].find("p").find("span").string = "{:.2%}".format(k1)
                    print(f"已修改 {key} 为：{new_value}")

        modified_html = str(soup.prettify())
        mail.content.html = modified_html

        return mail


processor_map = {
    "zhaochenxing@swhysc.com": CustomerAProcessor(),
}


def get_processor(email: str) -> ProcessorStrategy:
    """
    根据邮箱地址获取对应的处理器策略
    :param email: 邮箱地址
    :return: ProcessorStrategy 策略
    """
    return processor_map.get(email)
