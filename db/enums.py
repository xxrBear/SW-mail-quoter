import enum


class MailStateEnum(enum.Enum):
    """邮件处理状态"""

    UNPROCESSED = "unprocessed"  # 未处理
    PROCESSED = "processed"  # 已自动处理
    MANUAL = "manual"  # 人工处理
