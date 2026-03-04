import json
from datetime import datetime
from .base_exporter import BaseExporter
from utils.file_utils import generate_filename
from utils.logger import get_logger


class JSONExporter(BaseExporter):
    """
    Export processor insights to JSON format.
    """

    def __init__(self):
        self.logger = get_logger(self.__class__.__name__)

    def export(self, processor, output_path=None):

        self._validate_processor(processor)

        # Prepare data for JSON
        export_data = {
            "metadata": {
                "export_timestamp": datetime.now().isoformat(),
                "processor_type": processor.__class__.__name__,
                "processing_summary": processor.get_processing_summary(),
            },
            "insights": self._prepare_insights(processor.insights),
            "alerts": processor.alerts,
            "raw_data_summary": {
                "row_count": len(processor.data) if processor.data is not None else 0,
                "columns": (
                    list(processor.data.columns) if processor.data is not None else []
                ),
            },
        }

        # Generate filename if not provided
        if output_path is None:
            output_path = generate_filename(processor, ".json")

        # Ensure .json extension
        if not output_path.endswith(".json"):
            output_path += ".json"

        # Save to file
        try:
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(export_data, f, indent=2, default=str)

            print(f"✅ JSON report saved to: {output_path}")
            return output_path

        except Exception as e:
            print(f"❌ Error saving JSON file: {e}")
            raise

    def _prepare_insights(self, insights):
        """
        Prepare insights for JSON serialization.
        Converts pandas objects to JSON-serializable formats.
        """
        serializable_insights = {}

        for key, value in insights.items():
            # Handle pandas DataFrames and Series
            if hasattr(value, "to_dict"):
                try:
                    if hasattr(value, "index"):  # DataFrame or Series
                        serializable_insights[key] = value.to_dict()
                    else:
                        serializable_insights[key] = value
                except:
                    serializable_insights[key] = str(value)

            # Handle lists containing non-serializable objects
            elif isinstance(value, list):
                serializable_insights[key] = [
                    item.to_dict() if hasattr(item, "to_dict") else item
                    for item in value
                ]

            # Handle dictionaries
            elif isinstance(value, dict):
                serializable_insights[key] = {
                    k: v.to_dict() if hasattr(v, "to_dict") else v
                    for k, v in value.items()
                }

            # Handle other types
            else:
                try:
                    json.dumps(value)  # Test if serializable
                    serializable_insights[key] = value
                except:
                    serializable_insights[key] = str(value)

        return serializable_insights
