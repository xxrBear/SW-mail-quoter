import enum

from sqlalchemy import DateTime, Enum, String, func
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column


class Base(DeclarativeBase):
    pass


class MailStateEnum(enum.Enum):
    """邮件处理状态"""

    UNPROCESSED = "unprocessed"  # 未处理
    PROCESSED = "processed"  # 已自动处理
    MANUAL = "manual"  # 人工处理


class MailState(Base):
    __tablename__ = "mail_state"

    id: Mapped[int] = mapped_column(primary_key=True)
    created_time: Mapped[DateTime] = mapped_column(
        DateTime(timezone=True), server_default=func.now()
    )
    mail_hash: Mapped[str] = mapped_column(
        String(64), unique=True, index=True
    )  # 哈希值，唯一且加索引
    state: Mapped[MailStateEnum] = mapped_column(
        Enum(MailStateEnum, name="mail_state_enum"),
        default=MailStateEnum.UNPROCESSED,
        nullable=False,
    )
