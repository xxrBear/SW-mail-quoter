from contextlib import contextmanager
from db.engine import SessionLocal


@contextmanager
def session_scope():
    """提供一个自动提交、自动回滚、自动关闭的上下文"""
    session = SessionLocal()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()
