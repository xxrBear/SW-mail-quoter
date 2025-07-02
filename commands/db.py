import click

from db.setup import delete_row, drop_db, init_db


@click.command(name="db")
@click.option("-i", "--init-db", "do_init", is_flag=True, help="初始化数据库表结构")
@click.option("-d", "--drop-db", "do_drop", is_flag=True, help="删除数据库表结构")
@click.option(
    "-del",
    "--delete",
    "delete",
    type=click.IntRange(1, None),  # 限制为 >=1，无上限
    help="删除指定天数前的数据",
)
def cli_db(do_init, do_drop, delete):
    """执行数据库操作"""
    if do_init:
        init_db()
    elif do_drop:
        confirm = click.confirm("确定要删除所有数据库表？此操作不可恢复！")
        if not confirm:
            click.secho("操作取消", fg="yellow")
        drop_db()
    elif delete:
        count = delete_row(delete)
        click.secho(f"已删除{delete}天前的数据共{count}条", fg="green")
    else:
        click.secho("请输入参数 -i（初始化）或 -d（删除）", fg="yellow")
