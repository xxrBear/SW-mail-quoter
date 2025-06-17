from processor.base import ProcessorStrategy
from processor.customer.cbg import CustomerCBGProcessor

# 邮箱后缀与处理策略类的映射，只保存了一个实例(类似单例模式)
processor_map = {
    "swhysc.com": CustomerCBGProcessor(),  # 测试邮箱后缀
    "cgbchina.com.cn": CustomerCBGProcessor(),  # 广发银行
}


def get_processor(email: str) -> ProcessorStrategy:
    """
    根据邮箱地址后缀获取对应的处理策略类实例
    :param email: 邮箱地址
    :return: ProcessorStrategy 子类实例
    """
    email_suffix = email.split("@")[-1]

    return processor_map.get(email_suffix)


subject_sheet_map = {
    "看涨阶梯": "看涨阶梯",
    "看跌阶梯": "看跌阶梯",
    "二元看涨": "二元看涨",
}


def choose_sheet_by_subject(subject: str) -> str:
    """根据邮件主题选择对应的工作表名称
    :param subject: 邮件主题
    :return: 对应的工作表名称
    """

    for keyword, sheet_name in subject_sheet_map.items():
        if keyword in subject:
            return sheet_name
    return None
