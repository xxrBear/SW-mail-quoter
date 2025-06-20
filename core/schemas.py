from dataclasses import dataclass
from datetime import datetime
from email.message import Message
from typing import List, Optional, Union

from bs4 import BeautifulSoup


@dataclass
class MailContent:
    plain: str
    html: str
    attachments: Optional[Union[List[str], str]] = None  # noqa: F821
    nested: Optional[List["MailContent"]] = None


@dataclass
class EachMail:
    msg_id: str  # 邮件编号
    subject: str  # 邮件标题
    from_name: str  # 发件人名称
    from_addr: str  # 发件人邮箱地址
    content: MailContent
    message: Message  # 原始邮件内容
    sent_time: datetime  # 邮件发送时间，格式为 "YYYY-MM-DD HH:MM:SS"
    df_dict: dict  # 解析后的 DataFrame 字典类型数据
    soup: Optional[BeautifulSoup] = None  # BeautifulSoup 对象，解析后的邮件 HTML 内容
    sheet_name: Optional[str] = None
