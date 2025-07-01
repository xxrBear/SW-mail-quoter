import os
import pickle
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date

import xlwings as xw

from core.client import send_mail_client
from core.excel import ExcelHandler
from core.handler import MailHandler
from core.utils import print_banner, print_init_db, selected_excel_if_open
from db.models import MailState


def init_db():
    """初始化数据库和表结构表"""
    from sqlalchemy import inspect

    from db.engine import engine
    from db.models import Base

    inspector = inspect(engine)

    if inspector.has_table("mail_state"):
        print_banner("申万宏源处理脚本")
    else:
        Base.metadata.create_all(bind=engine)
        print_init_db("数据库表初始化完成......")


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
        mail_subjects = ExcelHandler.get_confirmed_mail_subject(sheet)
        mails = state.get_unprocessed_mails(sheet_name, mail_subjects)

        if not mails.count():
            return

        # 使用多线程发送邮件
        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = [
                executor.submit(send_mail_client.reply_mail, pickle.loads(p.mail_raw))
                for p in mails
            ]
            for f in as_completed(futures):
                try:
                    f.result()
                except Exception as e:
                    print(f"邮件发送失败: {e}")

        # 更新已处理邮件状态
        mail_ids = [m.id for m in mails]
        try:
            state.batch_update_mails_state(mail_ids)
        except Exception as e:
            print(f"更新数据库失败：{e}")
    finally:
        print_banner("邮件发送成功")
        if not run_in_background:
            wb.save()
        else:
            wb.save()
            wb.close()
            app.quit()


if __name__ == "__main__":
    # init_db()
    process_excel()
    # reply_emails("二元看涨")
