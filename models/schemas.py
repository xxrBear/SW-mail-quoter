from dataclasses import dataclass
from email.message import Message
from typing import List, Optional, Union


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
    from_name: str
    from_addr: str
    content: MailContent
    message: Message  # 原始邮件内容
