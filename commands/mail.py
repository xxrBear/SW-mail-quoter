import click

from core.handler import MailHandler


@click.group(name="mail")
def cli_mail():
    """邮件处理"""
    pass


@cli_mail.command("pull")
def pull():
    """拉取邮件并写入数据库"""
    MailHandler().pull_mails_to_db()
