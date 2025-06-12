from datetime import date

import xlwings as xw

from core.client import create_mail_client
from processor.registry import get_processor


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

        for mail in result_list:
            print(f"处理邮件: {mail.subject} 来自: {eamil_addr}")

            k1 = processor.process_excel(mail, wb)
            processed_mail = processor.process_mail_html(mail, k1)

            # 回复邮件
            mail_client.reply_mail(processed_mail)

    print("所有邮件处理完成，保存并关闭 Excel 文件...")
    wb.close()
    app.quit()


if __name__ == "__main__":
    process_excel_and_reply_mails()
