from functools import wraps

from db.engine import SessionLocal


def with_session(func):
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
