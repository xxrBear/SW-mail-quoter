from datetime import date

import pandas as pd
import xlwings as xw

from core.client import create_mail_client
from processor.base import choose_sheet_by_subject, get_processor


def process_excel_and_reply_mails():
    """处理 Excel 并回复邮件
    :param mail_client: 邮箱客户端实例
    :return: None
    """

    # 启动 Excel 应用
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open("./test.xlsm")
    except Exception as e:
        print("打开 Excel 文件失败:", e)
        app.quit()

    # 处理邮件并回复
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

            sheet_name = choose_sheet_by_subject(mail.subject)
            k1 = processor.process_excel(df, wb, sheet_name)
            processed_mail = processor.process_mail_html(mail, k1)

            # 回复邮件
            mail_client.reply_mail(processed_mail)
            break

    app.quit()


if __name__ == "__main__":
    process_excel_and_reply_mails()
