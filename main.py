import pickle
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date

import xlwings as xw

from core.client import send_mail_client
from core.handler import MailHandler
from core.utils import print_banner, print_init_db
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


def process_excel():
    """处理 Excel"""

    # 启动 Excel 应用
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open(r"./奇异期权.xlsm")
    except Exception as e:
        print("打开 Excel 文件失败:", e)
        app.quit()
        return

    # 处理邮件并回复
    try:
        mail_handler = MailHandler()
        mail_handler.handle(wb)

        print("所有邮件处理完成，保存并关闭 Excel 文件...")
    except:
        raise
    finally:
        wb.save()
        wb.close()
        app.quit()


def reply_emails(sheet_name: str):
    state = MailState()
    mails = state.get_unprocessed_mails(sheet_name)

    if not mails:
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
                print(f"发送失败: {e}")

    mail_ids = [m.id for m in mails]
    state.batch_update_mails_state(mail_ids)

    print_banner("邮件发送成功")


if __name__ == "__main__":
    # init_db()
    # process_excel()
    reply_emails("二元看涨")
