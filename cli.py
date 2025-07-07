import click

from commands.db import cli_db
from commands.proc import cli_process_excel
from commands.mail import cli_mail


@click.group()
def cli():
    """申万宏源报价处理命令行工具"""
    pass


cli.add_command(cli_db)
cli.add_command(cli_process_excel)
cli.add_command(cli_mail)


if __name__ == "__main__":
    cli()
