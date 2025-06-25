from datetime import date, timedelta

from sqlalchemy import DateTime, Enum, String, func
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column

from core.parser import get_mail_hash
from core.schemas import EachMail
from db.session import session_scope
from db.enums import MailStateEnum


class Base(DeclarativeBase):
    pass


class MailState(Base):
    __tablename__ = "mail_state"

    id: Mapped[int] = mapped_column(primary_key=True)

    created_time: Mapped[DateTime] = mapped_column(
        DateTime(timezone=True), server_default=func.now()
    )

    rev_time: Mapped[DateTime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), comment="询价时间"
    )

    subject: Mapped[str] = mapped_column(
        String(256), nullable=False, index=True, comment="邮件标题"
    )

    from_addr: Mapped[str] = mapped_column(
        String(256), nullable=False, comment="发件人", index=True
    )

    state: Mapped[MailStateEnum] = mapped_column(
        Enum(MailStateEnum, name="mail_state_enum"),
        default=MailStateEnum.UNPROCESSED,
        nullable=False,
        comment="邮件处理状态",
    )

    sheet_name: Mapped[str] = mapped_column(
        String(64), nullable=False, index=True, comment="excel 工作簿"
    )

    mail_hash: Mapped[str] = mapped_column(
        String(64), unique=True, index=True, comment="标题与发送时间组合的哈希值"
    )

    def update_or_create_record(self, mail: EachMail) -> None:
        """将处理结果更新或写入数据库"""
        mail_hash = get_mail_hash(mail)

        with session_scope() as session:
            mail_obj = (
                session.query(MailState)
                .filter_by(mail_hash=mail_hash, state=MailStateEnum.MANUAL)
                .one_or_none()
            )

            if mail_obj:
                mail_obj.state = MailStateEnum.PROCESSED
            else:
                mail_obj = MailState(
                    mail_hash=mail_hash,
                    sheet_name=mail.sheet_name,
                    state=MailStateEnum.PROCESSED,
                    subject=mail.subject,
                    from_addr=mail.from_addr,
                    rev_time=mail.sent_time,
                )
                session.add(mail_obj)
            session.commit()

    def mail_exists(self, mail: EachMail) -> bool:
        """检查邮件是否已存在"""
        mail_hash = get_mail_hash(mail)
        with session_scope() as session:
            c = (
                session.query(MailState).filter_by(mail_hash=mail_hash).first()
                is not None
            )
            return c

    def get_successful_mail_info(self) -> list:
        with session_scope() as session:
            mails = (
                session.query(MailState)
                .filter(
                    MailState.state == MailStateEnum.PROCESSED,
                    MailState.created_time >= date.today(),
                )
                .order_by(MailState.rev_time)
            )
            return [
                [
                    m.subject,
                    m.from_addr,
                    m.rev_time,
                    m.created_time + timedelta(hours=8),
                ]
                for m in mails
            ]
