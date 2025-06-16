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
    """在工作表中添加邮件主题字段"""
    sheet = wb.sheets[mail.sheet_name]
    last_cell = sheet.range("A100").end("up")

    if last_cell.value == "邮件标题":
        next_row = last_cell.row
    else:
        next_row = last_cell.row + 1
        sheet.range(f"A{next_row}").value = "邮件标题"
    sheet.range(f"{next_letter}{next_row}").value = mail.subject
    wb.save()
