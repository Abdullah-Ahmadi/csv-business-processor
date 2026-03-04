import logging

from abc import ABC, abstractmethod


class BaseProcessor(ABC):

    def __init__(self):
        self.data = None
        self.insights = {}
        self.alerts = []
        self.summary = {}
        self.logger = logging.getLogger(self.__class__.__name__)

    @abstractmethod
    def load_data(self):
        pass

    @abstractmethod
    def analyze(self):
        pass

    @abstractmethod
    def validate(self):
        if self.data is None:
            raise ValueError("No data loaded. Call load_data() first.")

        if self.data.empty:
            raise ValueError("Data is empty")

        return True

    def get_processing_summary(self):
        return {
            "rows_processed": len(self.data),
            "columns": ", ".join(list(self.data.columns)),
            "insights_count": len(self.insights),
            "alerts_count": len(self.alerts),
            "processing_mode": self.__class__.__name__.replace("Processor", ""),
        }

    def _standardize_columns_names(self):
        if self.data is not None:
            self.data.columns = [
                str(col).lower().replace(" ", "_") for col in self.data.columns
            ]
