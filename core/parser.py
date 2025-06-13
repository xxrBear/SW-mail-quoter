import base64
from email.header import decode_header
from email.message import Message
from email.utils import parseaddr
from typing import List, Optional, Tuple

import pandas as pd

from core.schemas import MailContent


def parse_html_to_dict(html: str) -> Optional[dict]:
    """
    解析邮件 HTML 的 table 内容，返回字典格式的数据
    :param html: 邮件 HTML 内容
    :return: 解析后的数据列表
    """
    try:
        df = pd.read_html(html)[0]  # 默认读第一张表
        df = df[[0, 1]]  # 确保只保留前两列
        df.dropna(subset=[0], inplace=True)  # 丢掉 key 是 NaN 的行

        result = {
            str(k).strip(): (None if pd.isna(v) else str(v).strip())
            for k, v in zip(df[0], df[1])
        }
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
        content_type = part.get_content_type()
        if content_type == "text/plain":
            plain += decode_part(part)
        elif content_type == "text/html":
            html += decode_part(part)
        elif part.get_filename():
            attachments.extend(parse_attachments(part))
        # 嵌套的multipart
        elif content_type.startswith("multipart/"):
            nested_multipart = extract_mail_content(part)
            nested.append(nested_multipart)

    return MailContent(plain, html, attachments, nested)


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
                    # 使用 base64 编码附件内容
                    "data": base64.b64encode(content).decode("ascii"),
                }
            )
    return attachments


def parse_from_info(msg: Message) -> Tuple[str, str]:
    """
    解析邮件的发件人信息（From 字段），返回发件人名称和邮箱地址

    支持对名称部分的编码解码处理，并去除多余的引号和尖括号，保证返回的名称和邮箱地址为干净的字符串

    :param msg: 邮件 Message 对象
    :return: (发件人名称, 发件人邮箱地址)
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

    subject = msg["Subject"]
    decoded_fragments = decode_header(subject)
    subject_str = ""
    for fragment, charset in decoded_fragments:
        if isinstance(fragment, bytes):
            subject_str += fragment.decode(charset or "utf-8", errors="replace")
        else:
            subject_str += fragment

    return subject_str


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
    cc.extend(filter_addresses(to_addrs, except_addresses))

    # 处理 CC 地址
    cc_addrs = msg.get("CC")
    cc.extend(filter_addresses(cc_addrs, except_addresses))

    return ",".join(cc)
