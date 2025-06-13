from functools import wraps

from db.engine import SessionLocal


def with_session(func):
    """装饰器，用于在函数执行前创建数据库会话，并在函数执行后提交或回滚事务"""

    @wraps(func)
    def wrapper(*args, **kwargs):
        session = SessionLocal()
        try:
            result = func(session, *args, **kwargs)
            session.commit()
            return result
        except:
            session.rollback()
            raise
        finally:
            session.close()

    return wrapper
