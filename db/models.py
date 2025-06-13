from sqlalchemy import DateTime, String, func
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column


class Base(DeclarativeBase):
    pass


class MailState(Base):
    __tablename__ = "mail_state"

    id: Mapped[int] = mapped_column(primary_key=True)
    created_time: Mapped[DateTime] = mapped_column(
        DateTime(timezone=True), server_default=func.now()
    )
    mail_hash: Mapped[str] = mapped_column(
        String(64), unique=True, index=True
    )  # 哈希值，唯一且加索引
