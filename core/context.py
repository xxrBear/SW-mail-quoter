from datetime import datetime


class AbnormalMailContext:
    """异常邮件对象"""

    def __init__(self):
        self.email = []
        self.hold_email = []

    def skip_mail(
        self,
        subject: str,
        sent_addr: str,
        sent_time: datetime,
        created_time: datetime,
        reason: str,
    ) -> None:
        """跳过异常邮件"""
        self.email.append(
            {
                "subject": subject,
                "reason": reason,
                "sent_addr": sent_addr,
                "sent_time": sent_time,
                "created_time": created_time,
            }
        )
        print(f"{reason}: {subject} 来自：{sent_addr}")

    def skip_hold_email(self, subject: str, sent_addr: str, reason: str) -> None:
        self.hold_email.append(
            {"subject": subject, "reason": reason, "sent_addr": sent_addr}
        )


mail_context = AbnormalMailContext()
