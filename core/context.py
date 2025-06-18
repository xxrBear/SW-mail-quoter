class AbnormalMailContext:
    """异常邮件对象"""

    def __init__(self):
        self.email = []

    def skip_mail(self, subject: str, sent_addr: str, reason: str) -> None:
        """跳过异常邮件"""
        self.email.append(
            {"subject": subject, "reason": reason, "sent_addr": sent_addr}
        )
        print(f"{reason}: {subject} 来自：{sent_addr}")


mail_context = AbnormalMailContext()
