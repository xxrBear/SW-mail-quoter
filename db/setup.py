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


def show_db():
    """展示当天的报价数据"""
    mails = MailState().get_db_info()

    doc1 = "\n\t".join([str(m) for m in mails if m.sheet_name == "二元看涨"])

    doc2 = "\n\t".join([str(m) for m in mails if m.sheet_name == "看涨阶梯"])

    doc = f"""
今日报价信息汇总
    
    处理成功：{len([m for m in mails if m.state == "UNPROCESSED"])} 条
    未处理：  {len([m for m in mails if m.state == "MAUNAL"])} 条
    手动处理：{len([m for m in mails if m.state == "PROCESSED"])} 条
    
    二元看涨报价邮件：
\t{doc1}
    
    看涨阶梯报价邮件：
\t{doc2}
    """
    return doc


def reset_row(_id: int):
    """通过 id 重置记录状态"""
    MailState().reset_state_by_id(_id)
