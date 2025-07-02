import click

from main import reply_emails


@click.command("reply")
@click.argument("sheet_name")
def cli_reply_emails(sheet_name):
    """回复指定邮件

    sheet_name: 待回复邮件类型"""
    reply_emails(sheet_name)
