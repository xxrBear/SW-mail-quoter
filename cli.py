import click

from commands.db import cli_db
from commands.reply import cli_reply_emails
from commands.proc import cli_process_excel


@click.group()
def cli():
    """申万宏源报价处理命令行工具"""
    pass


cli.add_command(cli_db)
cli.add_command(cli_reply_emails)
cli.add_command(cli_process_excel)


if __name__ == "__main__":
    cli()
