import json
import pickle
from datetime import date, datetime, timedelta, timezone
from typing import List, Optional

from sqlalchemy import JSON, DateTime, Enum, LargeBinary, String, func
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
        DateTime(timezone=True),
        default=datetime.now(timezone(timedelta(hours=8))),
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

    df_dict: Mapped[dict] = mapped_column(JSON, nullable=False, comment="邮件表格内容")

    soup: Mapped[str] = mapped_column(String(512), nullable=False, comment="邮件soup")

    def __repr__(self) -> str:
        return f"ID: {self.id:>3} 标题：{self.subject} 来自：<{self.from_addr}> 处理状态：{self.state.value}"

    def create_record(self, mail: EachMail) -> None:
        """将处理结果更新或写入数据库"""
        mail_hash = get_mail_hash(mail)

        with session_scope() as session:
            mail_obj = (
                session.query(MailState)
                .filter(MailState.mail_hash == mail_hash)
                .one_or_none()
            )

            if not mail_obj:
                mail_obj = MailState(
                    mail_hash=mail_hash,
                    sheet_name=mail.sheet_name,
                    subject=mail.subject,
                    from_addr=mail.from_addr,
                    rev_time=mail.sent_time,
                    underlying=mail.underlying,
                    mail_raw=pickle.dumps(mail),
                    df_dict=json.dumps(mail.df_dict),
                    soup=str(mail.soup),
                )
                session.add(mail_obj)

    def mail_exists(self, mail: EachMail) -> Optional["MailState"]:
        """检查邮件是否已存在"""
        mail_hash = get_mail_hash(mail)
        with session_scope() as session:
            session.expire_on_commit = False
            c = session.query(MailState).filter_by(mail_hash=mail_hash).first()
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
            return [[m.subject, m.from_addr, m.rev_time, m.created_time] for m in mails]

    def get_unprocessed_mails(
        self, sheet_name: str, mail_hash_list: list
    ) -> Optional["MailState"]:
        with session_scope() as session:
            mails = (
                session.query(MailState)
                .filter(
                    # MailState.rev_time >= date.today(),
                    MailState.state == MailStateEnum.UNPROCESSED,
                    MailState.sheet_name == sheet_name,
                    MailState.mail_hash.in_(mail_hash_list),
                )
                .order_by(MailState.rev_time)
            )
            return mails

    def batch_update_mails_state(self, mail_ids: list):
        with session_scope() as session:
            session.query(MailState).filter(MailState.id.in_(mail_ids)).update(
                {"state": MailStateEnum.PROCESSED}, synchronize_session="fetch"
            )

    def update_state_by_hash_mail(
        self, mail_hash: list, state: MailStateEnum = MailStateEnum.MANUAL
    ):
        with session_scope() as session:
            session.query(MailState).filter(MailState.mail_hash.in_(mail_hash)).update(
                {"state": state}, synchronize_session="fetch"
            )

    # ------------------------------------------------------------------------------------------
    # CLI 专用
    # ------------------------------------------------------------------------------------------
    def delete_records_older_than_days(self, days: int):
        """
        删除数据表中早于指定天数的记录

        :param days: 删除多少天前的数据（大于该天数的记录会被保留）
        :return: 删除的记录条数
        """
        with session_scope() as session:
            query = session.query(MailState).filter(
                MailState.created_time <= datetime.now() - timedelta(days=days)
            )
            count = query.count()
            query.delete()
        return count

    def clear_table(self) -> int:
        """删除数据表中所有的记录"""
        with session_scope() as session:
            total = session.query(MailState).count()
            session.query(MailState).delete()
            return total

    def get_today_unprocessed_mails(self) -> List["MailState"]:
        with session_scope() as session:
            session.expire_on_commit = False
            today = date.today()
            start_time = datetime.combine(today, datetime.min.time())
            mails = (
                session.query(MailState)
                .filter(
                    MailState.created_time >= start_time,
                    MailState.state == MailStateEnum.UNPROCESSED,
                )
                .order_by(MailState.rev_time)
                .all()
            )
            return mails

    def get_db_info(self):
        with session_scope() as session:
            today = date.today()
            session.expire_on_commit = False

            today = date.today()
            start_time = datetime.combine(today, datetime.min.time())
            mails = (
                session.query(MailState)
                .filter(MailState.created_time >= start_time)
                .all()
            )
            return mails

    def reset_state_by_id(self, _id):
        with session_scope() as session:
            obj = session.get(MailState, _id)
            if obj:
                obj.state = MailStateEnum.UNPROCESSED
