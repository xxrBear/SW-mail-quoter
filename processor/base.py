from abc import ABC, abstractmethod


class ProcessorStrategy(ABC):
    """处理器策略基类，定义了处理 Excel 和邮件 HTML 的抽象方法"""

    @abstractmethod
    def process_excel(self):
        pass

    @abstractmethod
    def process_mail_html(self):
        pass

    @abstractmethod
    def get_excel_processing_rules(self):
        pass

    @abstractmethod
    def cannot_quote(self) -> bool:
        pass
