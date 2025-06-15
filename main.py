from datetime import date

import xlwings as xw

from core.handler import MailHandler
from core.utils import print_init_db


def process_excel_and_reply_mails():
    """处理 Excel 并回复邮件"""

    # 启动 Excel 应用
    app = xw.App(visible=False, add_book=False)
    try:
        wb = app.books.open("./test.xlsm")
    except Exception as e:
        print("打开 Excel 文件失败:", e)
        app.quit()
        return

    # 处理邮件并回复
    try:
        mail_handler = MailHandler(since_date=date(2025, 5, 1))
        mail_handler.handle(wb)

        print("所有邮件处理完成，保存并关闭 Excel 文件...")
    except:
        raise
    finally:
        wb.save()
        wb.close()
        app.quit()


def init_db():
    """初始化数据库"""
    from db.engine import engine
    from db.models import Base

    Base.metadata.create_all(bind=engine)
    print_init_db("数据库初始化完成......")


if __name__ == "__main__":
    init_db()
    process_excel_and_reply_mails()
