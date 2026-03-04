from abc import ABC, abstractmethod


class BaseExporter(ABC):
    """Base class for other exporter classes."""

    @abstractmethod
    def export(self, processor, output_path=None):
        pass

    def _validate_processor(self, processor):
        """Validate processor has been properly analyzed."""
        if not processor.insights:
            raise ValueError("Processor has no insights. Run analyze() first.")
        return True
