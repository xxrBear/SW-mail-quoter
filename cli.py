import click

from db.setup import drop_db, init_db
from main import process_excel, reply_emails


@click.group()
def cli():
    """申万宏源报价处理命令行工具"""
    pass


@cli.command()
@click.option("-i", "--init-db", "do_init", is_flag=True, help="初始化数据库表结构")
@click.option("-d", "--drop-db", "do_drop", is_flag=True, help="删除数据库表结构")
def db(do_init, do_drop):
    """执行数据库操作"""
    if do_init:
        init_db()
    elif do_drop:
        confirm = click.confirm("确定要删除所有数据库表？此操作不可恢复！")
        if not confirm:
            click.secho("操作取消", fg="yellow")
        drop_db()
    else:
        click.secho("请输入参数 -i（初始化）或 -d（删除）", fg="yellow")


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


cli.add_command(db)
cli.add_command(cli_process_excel)
cli.add_command(cli_reply_emails)

if __name__ == "__main__":
    cli()
