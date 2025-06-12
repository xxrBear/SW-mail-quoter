import re

# -----------------------------------------------------------------
# 广发银行
# -----------------------------------------------------------------

# 指定报价字段
CBG_QUOTE_FIELD_MAPPING = {"看涨阶梯": "行权价格1（低）", "二元看涨": "行权价格"}


# 看涨阶梯对应 excel 表格的映射关系
CBG_BULL_LADDER_TUPLE = (
    "C23",
    {
        "挂钩标的合约": (
            "C3",
            lambda v: re.findall(r"[（(](.*?)[）)]", v)[0].replace(".", "").upper(),
        ),
        "产品启动日": ("C4", str),
        "交割日（双方资金清算日）": ("C5", str),
        "最低收益率（年化）": ("C9", str),
        "中间收益率（年化）": ("C10", str),
        "最高收益率（年化）": ("C11", str),
        "行权价格2（高）": ("C22", lambda v: v.replace("*", "")),
    },
)

# 二元看涨
CBG_BINARRY_CALL_TUPLE = (
    "C19",
    {
        "挂钩标的合约": (
            "C3",
            lambda v: re.findall(r"[（(](.*?)[）)]", v)[0].replace(".", "").upper(),
        ),
        "产品启动日": ("C4", str),
        "交割日（双方资金清算日）": ("C5", str),
        "最低收益率（年化）": ("C9", str),
        "最高收益率（年化）": ("C11", str),
        "期权费 （年化）": ("C8", str),
    },
)

# 不同表格名称对应不同处理样式的规则
CBG_EXCEL_PROCESSING_RULES_MAPPING = {
    "看涨阶梯": CBG_BULL_LADDER_TUPLE,
    "二元看涨": CBG_BINARRY_CALL_TUPLE,
}
