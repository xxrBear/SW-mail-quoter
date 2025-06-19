from datetime import date

from sqlalchemy import DateTime, Enum, String, func
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column

from core.parser import get_mail_hash
from core.schemas import EachMail
from db.decorator import with_session
from db.engine import SessionLocal
from db.enums import MailStateEnum


class Base(DeclarativeBase):
    pass


class MailState(Base):
    __tablename__ = "mail_state"

    id: Mapped[int] = mapped_column(primary_key=True)

    created_time: Mapped[DateTime] = mapped_column(
        DateTime(timezone=True), server_default=func.now()
    )

    subject: Mapped[str] = mapped_column(
        String(256), nullable=False, index=True, comment="邮件标题"
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

    @with_session
    def update_or_create_record(session: SessionLocal, self, mail: EachMail) -> None:
        """将处理结果更新或写入数据库"""
        mail_hash = get_mail_hash(mail)
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
            )
            session.add(mail_obj)
        session.commit()

    @with_session
    def mail_exists(session: SessionLocal, self, mail: EachMail) -> bool:
        """检查邮件是否已存在"""
        mail_hash = get_mail_hash(mail)
        return (
            session.query(MailState).filter_by(mail_hash=mail_hash).first() is not None
        )

    @with_session
    def count_today_sheet_names(
        session: SessionLocal, self, mail: EachMail
    ) -> MailStateEnum:
        """获取当天sheet_name对应的数量"""
        mail_count = (
            session.query(MailState)
            .filter(
                MailState.sheet_name == mail.sheet_name,
                MailState.state == MailStateEnum.PROCESSED,
                MailState.created_time >= date.today(),
            )
            .count()
        )

        return mail_count

    @with_session
    def get_successful_mail_info(session: SessionLocal, self) -> list:
        mails = session.query(MailState).filter(
            MailState.state == MailStateEnum.PROCESSED,
            MailState.created_time >= date.today(),
        )
        return [[m.subject, m.subject] for m in mails]
