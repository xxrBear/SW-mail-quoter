import email
import imaplib
import os
import smtplib
from collections import defaultdict
from datetime import date
from email.mime.message import MIMEMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import make_msgid
from typing import Dict, List, Union

from bs4 import BeautifulSoup

from core.context import mail_context
from core.parser import (
    gen_cc,
    parse_from_info,
    parse_html_to_dict,
    parse_mail_sent_time,
    parse_multipart_content,
    parse_subject,
)
from core.schemas import EachMail
from processor.registry import choose_sheet_by_subject


class EmailClient:
    def __init__(
        self,
        server: str,
        address: str,
        password: str,
        imap_port: int = 993,
        smtp_port: int = 465,
    ) -> None:
        self.server = server
        self.imap_port = imap_port
        self.smtp_port = smtp_port
        self.address = address
        self.password = password

    def connect(
        self, protocol: str = "imap"
    ) -> Union[imaplib.IMAP4_SSL, smtplib.SMTP_SSL]:
        """
        通用邮件客户端连接方法

        :param protocol: 协议类型，可选 "imap" 或 "smtp"
        :return: 返回对应的邮件客户端对象
        :raises: 连接失败时抛出相应异常
        """
        if protocol.lower() == "imap":
            client = imaplib.IMAP4_SSL(self.server, self.imap_port)
            error_type = imaplib.IMAP4.error
            protocol_name = "IMAP接收"
            port = self.imap_port
        elif protocol.lower() == "smtp":
            client = smtplib.SMTP_SSL(self.server, self.smtp_port)
            error_type = smtplib.SMTPException
            protocol_name = "SMTP发送"
            port = self.smtp_port
        else:
            raise ValueError(f"不支持的协议类型: {protocol}，请使用 'imap' 或 'smtp'")

        try:
            client.login(self.address, self.password)
        except error_type as e:
            error_msg = (
                f"服务器{self.server}:端口{port}, "
                f"邮箱地址{self.address}:密码{self.password}，"
                f"{protocol_name}登录失败: {e}"
            )
            print(error_msg)
            raise e

        return client

    def read_mail(
        self,
        folder: str = "INBOX",
        since_date: date = date.today(),
    ) -> Dict[str, List[EachMail]]:
        """读取所有邮件，并整理成字典

        :param folder: 邮件文件夹，默认为 "INBOX"
        :param since_date: 读取指定日期之后的邮件，默认为今天
        :return: 返回一个字典，键为发件人地址，值为 EachMail 对象列表
        """
        mail_client = self.connect(protocol="imap")

        mail_client.select(folder)  # type: ignore # 选择收件箱

        result_dict = defaultdict(list)

        # 根据条件搜索邮件（可选条件：ALL、UNSEEN、SUBJECT "关键字"）
        status, messages = mail_client.search(  # type: ignore
            None, "Since", since_date.strftime("%d-%b-%Y")
        )
        if status != "OK":
            print("未找到邮件")
            return result_dict

        # 邮件ID列表
        message_ids = messages[0].split()

        for msg_id in message_ids:
            # 邮件原始数据
            status, msg_data = mail_client.fetch(msg_id, "(RFC822)")  # type: ignore
            if status != "OK":
                continue

            # 邮件内容
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            # 邮件标题和发件人信息
            subject = parse_subject(msg)
            from_name, from_addr = parse_from_info(msg)

            # 筛选邮件
            if "衍生品交易" not in subject:
                continue

            filter_subject_list = ["看涨阶梯", "二元看涨"]
            if not any(i for i in filter_subject_list if i in subject):
                mail_context.skip_mail(
                    subject, from_addr, "邮件处理策略未配置，跳过邮件"
                )
                continue

            # 文本内容
            content = parse_multipart_content(msg)

            # HTML　表格的内容（字典类型）
            df_dict = parse_html_to_dict(content.html)
            if not df_dict:
                mail_context.skip_mail(subject, from_addr, "无可用表格内容，跳过邮件")
                continue

            sheet_name = choose_sheet_by_subject(subject)
            if not sheet_name:
                mail_context.skip_mail(
                    subject, from_addr, "未找到对应的工作表名称，跳过邮件"
                )
                continue

            soup = BeautifulSoup(content.html, "html.parser")

            sent_time = parse_mail_sent_time(msg)
            if not sent_time:
                mail_context.skip_mail(subject, from_addr, "无法解析发送时间，跳过邮件")
                continue

            result_dict[from_addr].append(
                EachMail(
                    msg_id=msg_id,
                    subject=subject,
                    from_name=from_name,
                    from_addr=from_addr,
                    content=content,
                    message=msg,
                    df_dict=df_dict,
                    soup=soup,
                    sheet_name=sheet_name,
                    sent_time=sent_time,
                )
            )

        mail_client.close()

        return result_dict

    def reply_mail(
        self,
        last_email: EachMail,
    ) -> None:
        """回复邮件"""

        reply_mime = self._build_reply_mime(last_email)

        self._send_reply_mail(reply_mime)

    def _build_reply_mime(self, last_email: EachMail) -> MIMEMultipart:
        """构建回复邮件的 MIMEMultipart 对象"""

        original_msg = last_email.message

        reply_mime = MIMEMultipart("mixed")
        reply_mime["Message-ID"] = make_msgid()
        reply_mime["In-Reply-To"] = original_msg["Message-ID"]
        reply_mime["References"] = original_msg["Message-ID"]

        # 构建邮件头
        reply_mime["From"] = self.address
        # reply_mime["To"] = original_msg["From"]
        reply_mime["To"] = "17855370672@163.com"

        reply_mime["Subject"] = f"Re: {original_msg['Subject']}"
        reply_mime["CC"] = gen_cc(original_msg, [self.address])

        # 构建回复邮件体
        reply_body = MIMEMultipart("related")
        reply_mime.attach(reply_body)
        reply_info = MIMEMultipart("alternative")

        if isinstance(last_email.content, MIMEMultipart):
            reply_info.attach(last_email.content)
        else:
            html_part = MIMEText(
                f"<p>{last_email.content.html}</p><br></b><hr/><b>以上由程序回复，以下是原始邮件：</b><hr/></b><br>",
                "html",
                "utf-8",
            )
            reply_info.attach(html_part)
        reply_body.attach(reply_info)
        reply_orig_message = MIMEMultipart("alternative")
        reply_orig_message.attach(original_msg)
        reply_body.attach(reply_orig_message)

        reply_body.attach(MIMEMessage(original_msg))
        return reply_mime

    def _send_reply_mail(self, reply_mime: MIMEMultipart) -> None:
        """发送回复邮件"""

        # SMTP客户端连接
        smtp_client = self.connect("smtp")

        try:
            smtp_client.send_message(reply_mime)
        except (smtplib.SMTPException, AttributeError, Exception) as e:
            print(f"邮件回复失败: {type(e).__name__}: {e}")
            raise
        finally:
            smtp_client.quit()


def create_mail_client():
    """从环境变量创建并返回邮件客户端实例"""
    required_env_vars = {
        "EMAIL_SMTP_SERVER": os.getenv("EMAIL_SMTP_SERVER"),
        "EMAIL_USER_NAME": os.getenv("EMAIL_USER_NAME"),
        "EMAIL_USER_PASS": os.getenv("EMAIL_USER_PASS"),
    }

    missing_vars = [key for key, value in required_env_vars.items() if not value]
    if missing_vars:
        raise RuntimeError(f"缺少必须的环境变量: {', '.join(missing_vars)}")

    return EmailClient(
        server=required_env_vars.get("EMAIL_SMTP_SERVER"),
        address=required_env_vars.get("EMAIL_USER_NAME"),
        password=required_env_vars.get("EMAIL_USER_PASS"),
    )


# 全局单例
mail_client = create_mail_client()
