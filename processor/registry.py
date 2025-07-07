from typing import Dict, Optional, Type, TypeVar

from processor.base import ProcessorStrategy
from processor.impl.cbg import CustomerCBGProcessor

# 邮箱后缀与处理策略类的映射，只保存了一个实例(类似单例模式)
processor_map: Dict[str, Type[ProcessorStrategy]] = {
    "swhysc.com": CustomerCBGProcessor,  # 测试邮箱后缀
    "cgbchina.com.cn": CustomerCBGProcessor,  # 广发银行
}

# 客户对应邮件抄送人
cc_map: Dict[str, str] = {
    "cgbchina.com.cn": "liunaiwei@swhysc.com,otc_sales_sh1@swhysc.com",
    "swhysc.com": "17855370672@163.com,zhaochengxin@swhysc.com",
}


def get_cc_map(email: str) -> Optional[str]:
    email_suffix = email.split("@")[-1]
    cc = cc_map.get(email_suffix)
    return cc


T = TypeVar("T", bound=ProcessorStrategy)


def get_processor(email: str) -> Optional[ProcessorStrategy]:
    """
    根据邮箱地址后缀获取对应的处理策略类实例
    :param email: 邮箱地址
    :return: ProcessorStrategy 子类实例
    """
    email_suffix = email.split("@")[-1]
    processor_cls = processor_map.get(email_suffix)
    if processor_cls is not None:
        return processor_cls()
    return None


subject_sheet_map = {
    "看涨阶梯": "看涨阶梯",
    "二元看涨": "二元看涨",
}


def choose_sheet_by_subject(subject: str) -> Optional[str]:
    """根据邮件主题选择对应的工作表名称
    :param subject: 邮件主题
    :return: 对应的工作表名称
    """

    for keyword, sheet_name in subject_sheet_map.items():
        if keyword in subject:
            return sheet_name
    return None
