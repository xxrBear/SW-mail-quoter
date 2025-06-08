import os
from datetime import date

import pandas as pd

from core.client import EmailClient
from processor.base import get_processor


def create_mail_client():
    """从环境变量创建并返回邮件客户端实例"""
    required_env_vars = {
        "EMAIL_SMTP_SERVER": os.getenv("EMAIL_SMTP_SERVER"),
        "EMAIL_USER_NAME": os.getenv("EMAIL_USER_NAME"),
        "EMAIL_USER_PASS": os.getenv("EMAIL_USER_PASS"),
    }

    missing_vars = [key for key, value in required_env_vars.items() if not value]
    if missing_vars:
        raise RuntimeError(f"缺少必须的环境变量: {', '.join(missing_vars)}")

    return EmailClient(
        server=required_env_vars.get("EMAIL_SMTP_SERVER"),
        address=required_env_vars.get("EMAIL_USER_NAME"),
        password=required_env_vars.get("EMAIL_USER_PASS"),
    )


def parse_mail_content(mail_client: EmailClient):
    """
    解析所有符合条件的邮件内容
    :param mail_client: 邮箱客户端实例
    :return: DataFrame
    """
    # 解析邮件内容
    result_dict = mail_client.read_mail(folder="INBOX", since_date=date(2025, 4, 21))

    for eamil_addr, result_list in result_dict.items():
        processor = get_processor(eamil_addr)
        if not processor:
            print(f"未找到处理器，邮箱地址: {eamil_addr}")
            continue

        last_mail = result_list[-1]
        print(f"处理邮件: {last_mail.subject} 来自: {eamil_addr}")

        df = pd.read_html(last_mail.content.html, index_col=0)[0]

        df.reset_index(inplace=True)
        k1 = processor.operate_excel(df)
        last_mail = processor.handle_mail_html(mail_client, last_mail, df, k1)

        mail_client.reply_mail(last_mail)


if __name__ == "__main__":
    mail_client = create_mail_client()

    parse_mail_content(mail_client)
