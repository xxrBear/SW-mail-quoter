from dataclasses import dataclass
from email.header import decode_header
from email.message import Message
from typing import Optional, Union, List, Tuple


@dataclass
class MailContent:
    plain: str
    html: str
    attachments: Optional[Union[List[str], str]] = None
    nested: Optional[List["MailContent"]] = None


def parse_alternative_content(msg: Message) -> MailContent:
    """解析multipart/alternative邮件内容"""

    # 获取边界标识
    boundary = msg.get_boundary()
    if not boundary:
        raise ValueError("Missing boundary in multipart/alternative")
    return parse_nested_multipart(msg)


def decode_part(part: Message) -> str:
    """解码单个邮件部分"""
    charset = part.get_content_charset("utf-8")

    payload = part.get_payload(decode=True)

    return payload.decode(charset, errors="replace")


def parse_attachments(part: Message) -> list:
    """解析附件内容"""
    attachments = []
    filename = part.get_filename()
    if filename:
        content = part.get_payload(decode=True)
        if content:
            attachments.append(
                {
                    "filename": filename,
                    "content_type": part.get_content_type(),
                    "size": len(content),
                    "data": content.hex(),  # 二进制转十六进制存储
                }
            )
    return attachments


def parse_nested_multipart(msg: Message) -> MailContent:
    """处理嵌套的multipart结构"""
    plain = ""
    html = ""
    attachments = []
    nested = []
    # 遍历所有子部分
    for part in msg.get_payload(decode=False):
        content_type = part.get_content_type()
        # 处理文本内容
        if content_type == "text/plain":
            plain += decode_part(part)
        elif content_type == "text/html":
            html += decode_part(part)
        elif part.get_filename():
            # 处理内嵌附件（如图片）
            attachments = parse_attachments(part)
            attachments.extend(attachments)
        elif content_type.startswith("multipart/"):
            # 递归处理嵌套结构
            nested_multipart = parse_nested_multipart(part)
            nested.append(nested_multipart)

    return MailContent(plain, html, attachments, nested)


def clean_name_addr(input: str) -> str:
    return input.replace('"', "").replace("<", "").replace(">", "")


def parse_from_info(msg: Message) -> Tuple[str, str]:
    from_info = decode_header(msg["From"])
    if len(from_info) == 1:
        from_info, charset = from_info[0]
        if isinstance(from_info, bytes):
            from_info = from_info.decode(charset or "utf-8")
        if " " in from_info:
            from_name, from_addr = from_info.split(" ")
        else:
            from_name = from_addr = from_info
    elif len(from_info) == 2:
        from_name, charset = from_info[0]
        if isinstance(from_name, bytes):
            from_name = from_name.decode(charset or "utf-8")
        from_addr, charset = from_info[1]
        if isinstance(from_addr, bytes):
            from_addr = from_addr.decode(charset or "utf-8")
    else:
        raise ValueError(f"From={from_info} is invalid")
    from_name, from_addr = clean_name_addr(from_name), clean_name_addr(from_addr)
    return from_name, from_addr


def parse_subject(msg: Message) -> str:
    subject, charset = decode_header(msg["Subject"])[0]
    if isinstance(subject, bytes):
        subject = subject.decode(charset or "utf-8")
    return subject


def gen_cc(msg: Message, except_addresses: List[str]) -> str:
    cc = []
    to_addrs = msg["To"]
    if "," in to_addrs:
        to_addrs = to_addrs.split(",")
        to_addrs = [
            addr
            for addr in to_addrs
            if not any(excep in addr for excep in except_addresses)
        ]
        cc.extend(to_addrs)
    else:
        if not any(excep in to_addrs for excep in except_addresses):
            cc.append(to_addrs)
    cc_addrs = msg["CC"]
    if cc_addrs:
        if "," in cc_addrs:
            cc_addrs = cc_addrs.split(",")
            cc_addrs = [
                addr
                for addr in cc_addrs
                if not any(excep in addr for excep in except_addresses)
            ]
            cc.extend(cc_addrs)
        else:
            if not any(excep in cc_addrs for excep in except_addresses):
                cc.append(cc_addrs)
    return ",".join(cc)
