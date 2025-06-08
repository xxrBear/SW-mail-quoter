import re

import pandas as pd
import xlwings as xw


class ProcessorStrategy:
    def operate_excel(self):
        raise NotImplementedError()


class CustomerAProcessor(ProcessorStrategy):
    def operate_excel(self, df: pd.DataFrame) -> float:
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


class CustomerBProcessor(ProcessorStrategy):
    def operate_excel(selfl): ...


processor_map = {
    "zhaochenxing@swhysc.com": CustomerAProcessor(),
}


def get_processor(email: str) -> ProcessorStrategy:
    """
    根据邮箱地址获取对应的处理器策略
    :param email: 邮箱地址
    :return: ProcessorStrategy 实例
    """
    return processor_map.get(email)
