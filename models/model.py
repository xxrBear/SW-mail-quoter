from datetime import date
from enum import Enum
from typing import List, Optional, Union

from pydantic import BaseModel


class EmailModel(BaseModel):
    to_receiver: Union[list, str]  # 收件人
    cc_receiver: Union[str, list] = []  # 抄送人
    subject: str  # 主题
    text: str  # 邮件内容
    files: Optional[list] = []  # 是否发送文件
    # server: Optional[str] = ""  # 邮件服务器
    username: Optional[str] = ""
    password: Optional[str] = ""

    def __repr__(self):
        return str(
            "{}-{}-{}-{}".format(
                self.to_receiver, self.cc_receiver, self.subject, self.text
            )
        )
