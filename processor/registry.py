from processor.base import ProcessorStrategy
from processor.customer.cbg import CustomerCBGProcessor

processor_map = {
    "zhaochenxing@swhysc.com": CustomerCBGProcessor(),
}


def get_processor(email: str) -> ProcessorStrategy:
    """
    根据邮箱地址获取对应的处理器策略
    :param email: 邮箱地址
    :return: ProcessorStrategy 策略
    """
    return processor_map.get(email)


subject_sheet_map = {
    "看涨阶梯": "看涨阶梯",
    "看跌阶梯": "看跌阶梯",
}


def choose_sheet_by_subject(subject: str) -> str:
    for keyword, sheet_name in subject_sheet_map.items():
        if keyword in subject:
            return sheet_name
    raise ValueError(f"未找到对应的工作表，主题: {subject}")
