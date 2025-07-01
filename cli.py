import click

from main import init_db, process_excel, reply_emails


@click.group()
def cli():
    """申万宏源报价处理命令行工具"""
    pass


@click.command(name="initdb")
def cli_init_db():
    """初始化数据库表"""
    init_db()


@click.command("proc")
def cli_process_excel():
    """处理报价Excel"""
    process_excel()


@click.command("reply")
@click.argument("sheet_name")
def cli_reply_emails(sheet_name):
    """回复指定邮件

    sheet_name: 待回复邮件类型"""
    reply_emails(sheet_name)


cli.add_command(cli_init_db)
cli.add_command(cli_process_excel)
cli.add_command(cli_reply_emails)

if __name__ == "__main__":
    cli()
