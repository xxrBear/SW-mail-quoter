import click

from main import process_excel


@click.command("proc")
def cli_process_excel():
    """处理报价Excel"""
    process_excel()
