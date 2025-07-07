import json
from collections import defaultdict
from types import SimpleNamespace

import click

from core.excel import ExcelHandler
from core.handler import MailHandler
from core.utils import print_banner
from db.models import MailState
from main import open_excel_with_filename, reply_emails
from processor.registry import get_processor, subject_sheet_map


@click.group(name="mail")
def cli_mail():
    """邮件处理"""
    pass


@cli_mail.command("pull")
def pull():
    """拉取邮件并写入数据库"""
    MailHandler().pull_quote_mails_to_db()


@cli_mail.command("reply")
@click.argument(
    "sheet_name", required=True, type=click.Choice(["二元看涨", "看涨阶梯"])
)
def cli_reply_emails(sheet_name):
    """回复指定邮件

    sheet_name: 待回复邮件类型"""
    reply_emails(sheet_name)


@cli_mail.command("proc")
def cli_proc_mail():
    """从数据库拉取邮件信息并写入 Excel 中"""
    wb, app, run_in_background = open_excel_with_filename()

    mails = MailState().get_today_unprocessed_mails()
    if not mails:
        return

    result_dict = defaultdict(list)
    for mail in mails:
        mail.df_dict = json.loads(mail.df_dict)
        mail.content = SimpleNamespace(html="")
        mail.sent_time = mail.rev_time
        result_dict[mail.from_addr].append(mail)

    # 处理未报价邮件并回复
    excel_handler = ExcelHandler()

    # 清空Sheet
    sheet_names = subject_sheet_map.keys()
    for _sheet_name in sheet_names:
        excel_handler.clear_sheet_columns(wb, _sheet_name)

    for email_addr, result_list in result_dict.items():
        print_banner("开始处理可报价邮件......")
        processor = get_processor(email_addr)  # 获取每个客户对应的邮件处理策略

        sheet_name_count_dict = {_sheet_name: 0 for _sheet_name in sheet_names}
        for mail in result_list:
            print(f"处理邮件: {mail.subject} 来自: 【{email_addr}】")

            # 处理 Excel 对应 Sheet
            excel_handler.copy_sheet_columns(
                wb, mail.sheet_name, sheet_name_count_dict[mail.sheet_name]
            )

            # 获取报价值，并写入待发送邮件内容中
            quote_value = processor.process_excel(
                mail, wb, sheet_name_count_dict[mail.sheet_name]
            )
            processor.process_mail_html(mail, quote_value)

            sheet_name_count_dict[mail.sheet_name] += 1

    print("所有邮件处理完成，保存并关闭 Excel 文件...")
    if not run_in_background:
        wb.save()
    else:
        wb.save()
        wb.close()
        app.quit()
