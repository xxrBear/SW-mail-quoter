from sqlalchemy import inspect

from core.utils import print_banner, print_init_db
from db.engine import engine
from db.models import Base, MailState


def init_db():
    """初始化数据库和表结构表"""

    inspector = inspect(engine)

    if inspector.has_table("mail_state"):
        print_banner("当前数据库表已完成初始化...")
    else:
        Base.metadata.create_all(bind=engine)
        print_init_db("数据库表初始化完成......")


def drop_db():
    """删除数据库所有表"""

    Base.metadata.drop_all(bind=engine)


def delete_row(days: int):
    """删除指定天数前的表数据"""
    return MailState().delete_records_older_than_days(days)


def clear_table():
    """删除表中所有数据"""
    return MailState().clear_table()
