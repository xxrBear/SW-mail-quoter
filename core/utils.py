import xlwings as xw

from core.schemas import EachMail


def print_banner(message: str, line_length: int = 120) -> None:
    line = "-" * line_length
    centered = f"{message.center(line_length)}"
    print(f"\n{line}\n{centered}\n{line}\n")


def print_init_db(message: str, line_length: int = 120) -> None:
    line = "*" * line_length
    centered = f"{message.center(line_length)}"
    print(f"\n{line}\n{centered}\n{line}\n")


def calc_next_letter(letter: str, count: int) -> str:
    """
    计算下一个字母
    :param letter: 当前字母
    :return: 下一个字母
    """
    finally_cell = str(chr(ord(letter) + count))
    return finally_cell


def add_excel_subject_cell(wb: xw.Book, mail: EachMail, next_letter: str) -> None:
    """在工作表中添加邮件标题字段"""
    sheet = wb.sheets[mail.sheet_name]
    value = "邮件标题"
    row, _ = find_position_in_column(sheet, value, "A")
    if not row:
        return

    sheet.range(f"{next_letter}{row}").value = mail.subject
    wb.save()


def find_position_in_column(sheet, keyword, col_index):
    """
    在指定工作表的某一列中查找包含 keyword 的单元格，返回其 (行号, 列号)
    """
    _100_row = 100  # 只处理一百行
    for row_index in range(1, _100_row + 1):
        cell = sheet.cells(row_index, col_index)
        value = cell.value
        if value is not None and str(value).strip() == str(keyword).strip():
            return row_index, col_index
    return None, None


def col_index_to_letter(n):
    """将列号（从 1 开始）转换为 Excel 列字母"""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def get_rate(underlying: str, value: float, wb: xw.Book) -> str:
    """
    根据输入值，返回对应区间的利率（百分比）
    """
    sheet = wb.sheets["标的价格"]

    thresholds = sheet.range("E1:I1").value

    if underlying.endswith("IDC"):
        rates = sheet.range("E2:I2").value
    elif underlying.startswith("AU") and underlying.endswith("SGE"):
        rates = sheet.range("E3:I3").value
    else:
        rates = sheet.range("E4:I4").value

    for threshold, rate in zip(thresholds, rates):
        if value <= threshold:
            return rate

    # 如果大于最大阈值，默认返回最低利率
    return rates[-1]


def get_risk_free_rate(underlying: str) -> str:
    if underlying.startswith("AU"):
        r = "2.4%"
    else:
        r = "4.5%"
    return r
