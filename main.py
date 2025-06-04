import imaplib
import smtplib
import email
from dataclasses import dataclass
from datetime import date
from email import policy

import os
from email.mime.message import MIMEMessage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.parser import BytesParser
from email.utils import make_msgid
from typing import List, Optional, Callable, Union

from util.mail_content_parser import (
    MailContent,
    parse_subject,
    parse_from_info,
    parse_alternative_content,
)


@dataclass
class EachMail:
    msg_id: str  # 邮件编号
    subject: str  # 邮件标题
    from_name: str
    from_addr: str
    content: MailContent


class EmailService:
    def __init__(
        self,
        server: str,
        address: str,
        password: str,
        imap_port: int = 993,
        smtp_port: int = 465,
        protocol: str = "imap",
    ) -> None:
        self.server = server
        self.imap_port = imap_port
        self.smtp_port = smtp_port
        self.address = address
        self.password = password
        self.protocol = protocol

    def connect(
        self, protocol: str = "imap"
    ) -> Union[imaplib.IMAP4_SSL, smtplib.SMTP_SSL]:
        """
        通用邮件客户端连接方法

        :param protocol: 协议类型，可选 "imap" 或 "smtp"
        :return: 返回对应的邮件客户端对象
        :raises: 连接失败时抛出相应异常
        """
        if self.protocol.lower() == "imap":
            client = imaplib.IMAP4_SSL(self.server, self.imap_port)
            error_type = imaplib.IMAP4.error
            protocol_name = "IMAP接收"
            port = self.imap_port
        elif self.protocol.lower() == "smtp":
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
        filter_func: Optional[Callable] = None,
    ) -> List[EachMail]:
        mail_client = self.connect(protocol="smtp")

        mail_client.select(folder)  # 选择收件箱

        result_list = []

        # 搜索所有邮件（可选条件：'ALL'/'UNSEEN'/'SUBJECT "关键词"'）
        status, messages = mail_client.search(
            None, "Since", since_date.strftime("%d-%b-%Y")
        )
        if status != "OK":
            print("未找到邮件")
            return result_list

        # 解析邮件ID列表
        message_ids = messages[0].split()

        # 遍历处理每封邮件
        for msg_id in message_ids:
            # 获取邮件原始数据
            status, msg_data = mail_client.fetch(msg_id, "(RFC822)")
            if status != "OK":
                continue

            # 解析邮件内容
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)
            # print(msg)

            # 解码邮件头信息
            subject = parse_subject(msg)
            from_name, from_addr = parse_from_info(msg)
            if filter_func is not None and not filter_func(
                from_addr, subject
            ):  # 筛选邮件，不满足直接跳过
                continue

            # 文本内容
            content = parse_alternative_content(msg)
            result_list.append(EachMail(msg_id, subject, from_name, from_addr, content))
        mail_client.close()
        return result_list

    def reply_mail(
        self,
        msg_id: str,
        reply_content: Union[str, MIMEMultipart],
        folder: str = "INBOX",
    ) -> None:
        rcv_client = self.rcv_connect()
        try:
            rcv_client.select(folder)  # 选择收件箱
        except imaplib.IMAP4.abort as e:
            rcv_client = self.rcv_connect()
            rcv_client.select(folder)
        status, msg_data = rcv_client.fetch(msg_id, "(RFC822)")
        if status != "OK":
            raise ValueError(f"MsgID={msg_id} is not in folder={folder}!")
        rcv_client.close()
        raw_email = msg_data[0][1]
        original_msg = BytesParser(policy=policy.default).parsebytes(raw_email)
        # original = email.message_from_string(raw_email.decode('utf-8'))
        # 开始回复
        reply_msg = MIMEMultipart("mixed")
        reply_msg["Message-ID"] = make_msgid()
        reply_msg["In-Reply-To"] = original_msg["Message-ID"]
        reply_msg["References"] = original_msg["Message-ID"]
        # 构建邮件头
        reply_msg["From"] = self.address
        reply_msg["To"] = original_msg["From"]
        reply_msg["Subject"] = f"Re: {original_msg['Subject']}"
        reply_msg["CC"] = gen_cc(original_msg, [self.address])
        # 构建回复邮件体
        reply_body = MIMEMultipart("related")
        reply_msg.attach(reply_body)
        reply_info = MIMEMultipart("alternative")
        if isinstance(reply_content, MIMEMultipart):
            reply_info.attach(reply_content)
        else:
            html_part = MIMEText(
                f"<p>{reply_content}</p><br></b><hr/><b>以上由程序回复，以下是原始邮件：</b><hr/></b><br>",
                "html",
                "utf-8",
            )
            reply_info.attach(html_part)
        reply_body.attach(reply_info)
        reply_orig_message = MIMEMultipart("alternative")
        reply_orig_message.attach(original_msg)
        reply_body.attach(reply_orig_message)

        reply_body.attach(MIMEMessage(original_msg))

        # 开始回复
        snd_client = self.snd_connect()
        try:
            snd_client.send_message(reply_msg)
        except smtplib.SMTPException as e:
            print(f"SMTP错误 [{e}]")
        except AttributeError as e:
            print(f"参数类型错误: {str(e)}")
            raise e
        except Exception as e:
            print(f"未知错误: {str(e)}")
        else:
            snd_client.close()


if __name__ == "__main__":
    # 获取环境变量
    EMAIL_SMTP_SERVER = os.getenv("EMAIL_SMTP_SERVER")
    EMAIL_USER_NAME = os.getenv("EMAIL_USER_NAME")
    EMAIL_USER_PASS = os.getenv("EMAIL_USER_PASS")

    if not all((EMAIL_SMTP_SERVER, EMAIL_USER_NAME, EMAIL_USER_PASS)):
        raise RuntimeError(
            f"缺少必须参数 EMAIL_USER_PASS {EMAIL_USER_PASS} EMAIL_USER_NAME {EMAIL_USER_NAME} EMAIL_USER_SERVER {EMAIL_SMTP_SERVER}"
        )

    folder = "INBOX"

    mail_client = EmailService(
        server=EMAIL_SMTP_SERVER, address=EMAIL_USER_NAME, password=EMAIL_USER_PASS
    )

    # def filter_mail(from_addr: str, subject: str) -> bool:  # 需要定义一个函数来筛选要处理的邮件
    #     if '衍生品交易' in subject and 'liunaiwei' in from_addr:
    #         return True
    #     return False

    result_list = mail_client.read_mail(folder=folder, since_date=date(2025, 4, 21))
    last_mail: EachMail = result_list[-2]

    import pandas as pd

    df: pd.DataFrame = pd.read_html(last_mail.content.html, index_col=0)[0]
    df.reset_index(inplace=True)
    # df.iloc[16, 1] = 110
    # print(df.values[2])
    # print(df.values.flatten())
    # print(df.stack().to_list())
    # for column_name, column_values in df.iterrows():
    # print(column_values.values)

    import xlwings as xw

    # 启动 Excel 应用
    app = xw.App(visible=False, add_book=False)

    wb = app.books.open("./test.xlsm")

    # 尝试常规关闭
    try:
        wb = app.books.open("./test.xlsm")
        sheet = wb.sheets[0]
        sheet.range("A18").value = "hello world3"
        wb.save()
    except Exception as e:
        print("操作 Excel 失败：", e)
    finally:
        wb.close()
        app.quit()
