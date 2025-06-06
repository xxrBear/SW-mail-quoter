import os
from datetime import date

import pandas as pd

from core.client import EmailClient
from models.schemas import EachMail


def create_mail_client():
    # 获取环境变量
    EMAIL_SMTP_SERVER = os.getenv("EMAIL_SMTP_SERVER")
    EMAIL_USER_NAME = os.getenv("EMAIL_USER_NAME")
    EMAIL_USER_PASS = os.getenv("EMAIL_USER_PASS")

    # 参数检查
    if not all((EMAIL_SMTP_SERVER, EMAIL_USER_NAME, EMAIL_USER_PASS)):
        raise RuntimeError(
            f"缺少必须参数 EMAIL_USER_PASS {EMAIL_USER_PASS} EMAIL_USER_NAME {EMAIL_USER_NAME} EMAIL_USER_SERVER {EMAIL_SMTP_SERVER}"
        )

    # 实例化邮箱客户端
    mail_client = EmailClient(
        server=EMAIL_SMTP_SERVER, address=EMAIL_USER_NAME, password=EMAIL_USER_PASS
    )
    return mail_client


def parse_mail_content(mail_client: EmailClient):
    """
    解析邮件内容
    :param mail_client: 邮箱客户端实例
    :return: DataFrame
    """
    # 解析邮件内容
    result_list = mail_client.read_mail(folder="INBOX", since_date=date(2025, 4, 21))
    # print(result_list)
    last_mail: EachMail = result_list[-2]

    df: pd.DataFrame = pd.read_html(last_mail.content.html, index_col=0)[0]
    # print(last_mail.content.html)
    # print(last_mail.content.plain)

    df.reset_index(inplace=True)
    return df, last_mail


def operate_excel(df: pd.DataFrame) -> float:
    """
    操作 Excel 文件
    :param df: 解析后的 DataFrame
    :return: k1: 从 Excel 中获取的值
    """
    # 这里可以添加对 DataFrame 的处理逻辑
    import re

    import xlwings as xw

    # 启动 Excel 应用
    app = xw.App(visible=False, add_book=False)

    wb = app.books.open("./test.xlsm")

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


def handle_mail_html(
    mail_client: EmailClient, last_mail: EachMail, df: pd.DataFrame, k1: float
):
    """
    处理邮件 HTML 内容
    :param mail_client: 邮箱客户端实例
    :param last_mail: 最后一封邮件实例
    :param k1: 从 Excel 中获取的值
    :return: None
    """
    from bs4 import BeautifulSoup

    html_content = last_mail.content.html
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
    last_mail.content.html = modified_html
    mail_client.reply_mail(last_mail.msg_id, last_mail.content)


if __name__ == "__main__":
    # def filter_mail(from_addr: str, subject: str) -> bool:  # 需要定义一个函数来筛选要处理的邮件
    #     if '衍生品交易' in subject and 'liunaiwei' in from_addr:
    #         return True
    #     return False

    mail_client = create_mail_client()

    df, last_email = parse_mail_content(mail_client)
    k1 = operate_excel(df)

    handle_mail_html(mail_client, last_email, df, k1)

    # 操作邮件 HTML 模块
    # print(k1)
    # print(last_mail.content.html)
    # mail_client.reply_mail(last_mail.msg_id, last_mail.content)
