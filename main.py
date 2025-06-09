from datetime import date

import pandas as pd

from core.client import create_mail_client
from processor.base import get_processor


def process_and_reply_mails():
    """处理并回复邮件
    :param mail_client: 邮箱客户端实例
    :return: None
    """

    mail_client = create_mail_client()
    result_dict = mail_client.read_mail(folder="INBOX", since_date=date(2025, 4, 21))

    for eamil_addr, result_list in result_dict.items():
        processor = get_processor(eamil_addr)
        if not processor:
            print(f"未找到对应的邮箱处理策略，邮箱地址: {eamil_addr}")
            continue

        for mail in result_list:
            print(f"处理邮件: {mail.subject} 来自: {eamil_addr}")

            df = pd.read_html(mail.content.html, index_col=0)[0]
            # print(df) # DataFrame
            df.reset_index(inplace=True)

            k1 = processor.operate_excel(df)
            processed_mail = processor.process_mail_html(mail_client, mail, df, k1)

            # 回复邮件
            mail_client.reply_mail(processed_mail)
            break


if __name__ == "__main__":
    process_and_reply_mails()
