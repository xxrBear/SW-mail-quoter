import base64
import hashlib
from datetime import datetime
from email.header import decode_header
from email.message import Message
from email.utils import parseaddr, parsedate_to_datetime
from typing import List, Optional, Tuple

from bs4 import BeautifulSoup

from core.schemas import EachMail, MailContent


def get_mail_hash(mail: EachMail) -> str:
    """
    生成邮件的唯一哈希值，用于标识邮件
    """
    subject = mail.subject

    sent_time = mail.sent_time

    join_str = f"{subject} - {sent_time.strftime('%Y-%m-%d %H:%M:%S')}"
    hash_obj = hashlib.sha256(join_str.encode("utf-8"))
    return hash_obj.hexdigest()


def parse_mail_sent_time(msg: Message) -> Optional[datetime]:
    """
    解析邮件的发送时间，返回格式化后的字符串
    """
    try:
        date_str = msg["Date"]
        if not date_str:
            return None

        # 解析邮件日期
        dt = parsedate_to_datetime(date_str)
        return dt if dt else None
    except Exception as e:
        print(f"解析发送时间失败: {e}")
        return None


def parse_html_to_dict(html: str) -> Optional[dict]:
    """
    解析邮件 HTML 的 table 内容，返回字典格式的数据
    """
    try:
        soup = BeautifulSoup(html, "html.parser")
        table = soup.find("table")
        if not table:
            return None

        result = {}
        for row in table.find_all("tr"):
            cols = row.find_all(["td", "th"])
            if len(cols) >= 2:
                key = cols[0].get_text(strip=True)
                value = cols[1].get_text(strip=True)
                if key:  # 丢掉 key 为 "" 的行
                    result[key] = value if value else None

        # print(f"BeautifulSoup 解析的结果：{result}")

        return result
    except Exception as e:
        print(f"解析失败: {e}")
        return None


def parse_multipart_content(msg: Message) -> MailContent:
    """解析 multipart 类型的邮件"""
    if not msg.is_multipart():
        raise ValueError("邮件不是 multipart 类型")

    return extract_mail_content(msg)


def extract_mail_content(msg: Message) -> MailContent:
    """解析邮件内容"""
    plain = ""
    html = ""
    attachments = []
    nested = []

    payload = msg.get_payload(decode=False)  # 获取邮件的内容部分

    for part in payload:
        content_type = part.get_content_type()  # type: ignore
        if content_type == "text/plain":
            plain += decode_part(part)  # type: ignore
        elif content_type == "text/html":
            html += decode_part(part)  # type: ignore
        elif part.get_filename():  # type: ignore
            attachments.extend(parse_attachments(part))  # type: ignore
        # 嵌套的multipart
        elif content_type.startswith("multipart/"):
            nested_multipart = extract_mail_content(part)  # type: ignore
            nested.append(nested_multipart)

    return MailContent(plain, html, attachments, nested)


def decode_part(part: Message) -> str:
    """解码单个邮件部分"""
    charset = part.get_content_charset("utf-8")

    payload = part.get_payload(decode=True)

    if isinstance(payload, bytes):
        return payload.decode(charset, errors="replace")
    elif isinstance(payload, str):
        return payload
    else:
        return ""


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
                    # 使用 base64 编码附件内容
                    "data": base64.b64encode(content).decode("ascii"),  # type: ignore
                }
            )
    return attachments


def parse_from_info(msg: Message) -> Tuple[str, str]:
    """
    解析邮件的发件人信息（From 字段），返回发件人名称和邮箱地址
    """

    raw_from = msg["From"]
    name, addr = parseaddr(raw_from)
    if name:
        decoded_name, charset = decode_header(name)[0]
        if isinstance(decoded_name, bytes):
            name = decoded_name.decode(charset or "utf-8")
        else:
            name = decoded_name
    # 去除多余符号
    name = name.replace('"', "").strip() if name else ""
    addr = addr.replace("<", "").replace(">", "").strip()
    # print(f"发送人: 姓名：{name} 邮箱：<{addr}>")

    return name, addr


def parse_subject(msg: Message) -> str:
    """解析邮件主题 Subject 字段 ，并处理可能的编码"""

    subject = ""
    for fragment, charset in decode_header(msg["Subject"]):
        if isinstance(fragment, bytes):
            # 如果 charset 不存在或非法，就兜底 utf-8
            safe_charset = (charset or "utf-8").lower()
            if safe_charset in ("unknown-8bit", "x-unknown"):
                safe_charset = "utf-8"
            try:
                fragment = fragment.decode(safe_charset, errors="replace")
            except LookupError:
                fragment = fragment.decode("utf-8", errors="replace")
        subject += fragment
    return subject


def filter_addresses(addresses: str, except_addresses: List[str]) -> List[str]:
    """过滤地址字符串，排除包含except_addresses中任意关键字的地址"""
    if not addresses:
        return []
    addr_list = [addr.strip() for addr in addresses.split(",")]
    filtered = [
        addr
        for addr in addr_list
        if not any(excep in addr for excep in except_addresses)
    ]
    return filtered


def gen_cc(msg: Message, except_addresses: List[str]) -> str:
    """生成过滤后的收件人和抄送人列表（逗号分隔字符串）"""
    cc = []

    # 处理 To 地址
    to_addrs = msg.get("To")
    cc.extend(filter_addresses(to_addrs or "", except_addresses))

    # 处理 CC 地址
    cc_addrs = msg.get("CC")
    cc.extend(filter_addresses(cc_addrs or "", except_addresses))

    return ",".join(cc)
