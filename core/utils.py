def print_banner(message: str, line_length: int = 120) -> None:
    line = "-" * line_length
    centered = f"{message.center(line_length)}"
    print(f"\n{line}\n{centered}\n{line}\n")


def print_init_db(message: str, line_length: int = 120) -> None:
    line = "*" * line_length
    centered = f"{message.center(line_length)}"
    print(f"\n{line}\n{centered}\n{line}\n")
