import re


class BaseHandler:
    def __init__(self, quote_name: str, quote_line: int):
        self.quote_name = quote_name  # 报价字段名称
        self.quote_line = quote_line  # 报价段所对应 sheet 的行


# --------------------------------------------------------------------------------------------------
# 广发银行
# --------------------------------------------------------------------------------------------------


class CBGBullLadderHandler(BaseHandler):
    """广发银行看涨阶梯处理类"""

    @property
    def fields_rule_dict(self) -> dict:
        return {
            "挂钩标的合约": (
                "3",
                lambda v: re.findall(r"[（(](.*?)[）)]", v)[0].replace(".", "").upper(),
            ),
            "产品启动日": ("4", str),
            "交割日（双方资金清算日）": ("5", str),
            "最低收益率（年化）": ("9", str),
            "中间收益率（年化）": ("10", str),
            "最高收益率（年化）": ("11", str),
            "行权价格2（高）": ("22", lambda v: v.replace("*", "")),
            "期权费（年化）": ("8", str),
        }

    @property
    def other_dict(self) -> dict:
        return {"交易日": "14", "VOL": "12", "标的合约": "3", "无风险利率": "17"}


class CBGBinarryCallHandler(BaseHandler):
    """广发银行二元看涨处理类"""

    @property
    def fields_rule_dict(self) -> dict:
        return {
            "挂钩标的合约": (
                "3",
                lambda v: re.findall(r"[（(](.*?)[）)]", v)[0].replace(".", "").upper(),
            ),
            "产品启动日": ("4", str),
            "交割日（双方资金清算日）": ("5", str),
            "最低收益率（年化）": ("9", str),
            "最高收益率（年化）": ("11", str),
            "期权费 （年化）": ("8", str),
        }

    @property
    def other_dict(self) -> dict:
        return {"交易日": "13", "VOL": "11", "标的合约": "3", "无风险利率": "16"}


CBG_SHEET_HANDLER = {
    "看涨阶梯": CBGBullLadderHandler("行权价格1（低）", 23),
    "二元看涨": CBGBinarryCallHandler("行权价格", 19),
}


def get_sheet_handler(sheet_name: str) -> BaseHandler:
    return CBG_SHEET_HANDLER.get(sheet_name)


# --------------------------------------------------------------------------------------------------
# 广发银行
# --------------------------------------------------------------------------------------------------
