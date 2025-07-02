import pickle
from datetime import date, timedelta
from typing import Optional

from sqlalchemy import DateTime, Enum, LargeBinary, String, func
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column

from core.parser import get_mail_hash
from core.schemas import EachMail
from db.enums import MailStateEnum
from db.session import session_scope


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

    underlying: Mapped[str] = mapped_column(
        String(256), nullable=False, comment="标的合约"
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

    mail_raw: Mapped[bytes] = mapped_column(LargeBinary, comment="邮件内容")

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
                    subject=mail.subject,
                    from_addr=mail.from_addr,
                    rev_time=mail.sent_time,
                    underlying=mail.underlying,
                    mail_raw=pickle.dumps(mail),
                )
                session.add(mail_obj)

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

    def get_unprocessed_mails(
        self, sheet_name: str, mail_subjects: list
    ) -> Optional["MailState"]:
        with session_scope() as session:
            mails = (
                session.query(MailState)
                .filter(
                    # MailState.rev_time >= date.today(),
                    MailState.state == MailStateEnum.UNPROCESSED,
                    MailState.sheet_name == sheet_name,
                    MailState.subject.in_(mail_subjects),
                )
                .order_by(MailState.rev_time)
            )
            return mails

    def batch_update_mails_state(self, mail_ids: list):
        with session_scope() as session:
            session.query(MailState).filter(MailState.id.in_(mail_ids)).update(
                {"state": MailStateEnum.PROCESSED}, synchronize_session="fetch"
            )
