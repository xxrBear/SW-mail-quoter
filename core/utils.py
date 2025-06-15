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
