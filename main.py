import os
import pickle
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date

import xlwings as xw

from core.client import send_mail_client
from core.excel import ExcelHandler
from core.handler import MailHandler
from core.utils import print_banner, selected_excel_if_open
from db.models import MailState
from processor.registry import get_processor


def open_excel_with_filename():
    """
    启动 Excel 应用并打开指定文件，失败时自动关闭 Excel
    """
    filename = os.getenv("EXCEL_FILENAME")
    wb, app = selected_excel_if_open(filename)

    if wb and app:
        return wb, app, False

    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(filename)
        return wb, app, True
    except Exception as e:
        print(f"无法打开当前 Excel {filename}")
        raise e


def process_excel():
    """处理 Excel"""
    wb, app, run_in_background = open_excel_with_filename()

    # 处理邮件并回复
    try:
        mail_handler = MailHandler()
        mail_handler.handle(wb)

    except:
        raise
    finally:
        print("所有邮件处理完成，保存并关闭 Excel 文件...")
        if not run_in_background:
            wb.save()
        else:
            wb.save()
            wb.close()
            app.quit()


def reply_emails(sheet_name: str):
    """回复邮件"""
    wb, app, run_in_background = open_excel_with_filename()

    try:
        state = MailState()
        sheet = wb.sheets[sheet_name]
        mail_hash_dict = ExcelHandler.get_confirmed_mail_hash_and_price(sheet)

        mails = state.get_unprocessed_mails(sheet_name, mail_hash_dict.keys())

        # 再次修改邮件的报价值，因为可能人为修改
        send_dict = {}
        confirmed_hash_list = []
        for m in mails:
            processor = get_processor(m.from_addr)
            mail_raw = pickle.loads(m.mail_raw)
            processor.process_mail_html(mail_raw, mail_hash_dict.get(m.mail_hash))
            send_dict[m.id] = mail_raw
            confirmed_hash_list.append(m.mail_hash)

        successful_ids = []
        # 使用多线程发送邮件
        if send_dict:
            with ThreadPoolExecutor(max_workers=10) as executor:
                future_map = {
                    executor.submit(send_mail_client.reply_mail, raw): id
                    for id, raw in send_dict.items()
                }
                for f in as_completed(future_map):
                    mail_id = future_map[f]
                    try:
                        f.result()
                        successful_ids.append(mail_id)
                    except Exception as e:
                        print(f"邮件发送失败: {e}")

        # 更新已处理邮件状态
        if successful_ids:
            try:
                state.batch_update_mails_state(successful_ids)
            except Exception as e:
                print(f"更新数据库失败：{e}")

        # 写入今日成功报价数据
        try:
            ExcelHandler().process_successful_mails_sheet(wb)
        except Exception as e:
            print(f"写入今日成功报价报错：{e}")

        # 写入被业务人员拒绝的数据
        reject_hash_list = ExcelHandler.get_reject_mail_hash(sheet)
        if reject_hash_list:
            MailState().update_state_by_hash_mail(reject_hash_list)

        print_banner("邮件发送成功")
    finally:
        if not run_in_background:
            wb.save()
        else:
            wb.save()
            wb.close()
            app.quit()


if __name__ == "__main__":
    # init_db()
    # process_excel()
    reply_emails("看涨阶梯")
