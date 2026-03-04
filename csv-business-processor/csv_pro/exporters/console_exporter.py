from .base_exporter import BaseExporter
from utils.formatters import format_currency, format_percentage
from utils.logger import get_logger


class ConsoleExporter(BaseExporter):
    """
    Export processor insights to the console.
    """

    def __init__(self):
        self.logger = get_logger(self.__class__.__name__)

    def export(self, processor, output_path=None):

        self._validate_processor(processor)

        output_lines = []

        # Header
        output_lines.append("\n" + "=" * 60)
        output_lines.append(
            f"📊 {processor.__class__.__name__.replace('Processor', '')} Analysis Report"
        )
        output_lines.append("=" * 60)

        # Summary
        output_lines.append("\n📈 SUMMARY")
        output_lines.append("-" * 40)

        summary = processor.get_processing_summary()
        for key, value in summary.items():
            output_lines.append(f"  • {key.replace('_', ' ').title()}: {value}")

        # Key insights
        output_lines.append("\n🔍 KEY INSIGHTS")
        output_lines.append("-" * 40)

        for key, value in processor.insights.items():
            if isinstance(value, (int, float)):
                # Format financial metrics
                if any(
                    term in key.lower()
                    for term in ["revenue", "amount", "cost", "price", "spent", "value"]
                ):
                    output_lines.append(
                        f"  • {self._format_key(key)}: {format_currency(value)}"
                    )
                elif (
                    "percentage" in key.lower()
                    or "rate" in key.lower()
                    or "growth" in key.lower()
                ):
                    output_lines.append(
                        f"  • {self._format_key(key)}: {format_percentage(value)}"
                    )
                else:
                    output_lines.append(f"  • {self._format_key(key)}: {value:,}")
            elif isinstance(value, dict) and len(value) <= 3:
                items = ", ".join(
                    f"{k}: {self._format_value(v)}" for k, v in value.items()
                )
                output_lines.append(f"  • {self._format_key(key)}: {items}")
            elif not isinstance(value, (list, dict)):
                output_lines.append(f"  • {self._format_key(key)}: {value}")

        # Display alerts
        if processor.alerts:
            output_lines.append("\n🚨 ALERTS & RECOMMENDATIONS")
            output_lines.append("-" * 40)

            for i, alert in enumerate(processor.alerts[:5], 1):
                icon = (
                    "⚠️  "
                    if "urgent" in alert.lower() or "critical" in alert.lower()
                    else "• "
                )
                output_lines.append(f"  {icon}{alert}")

            if len(processor.alerts) > 5:
                output_lines.append(
                    f"  ... and {len(processor.alerts) - 5} more alerts"
                )

        # Top Products
        if "top_products" in processor.insights:
            output_lines.append("\n🏆 TOP PERFORMING ITEMS")
            output_lines.append("-" * 40)

            top_products = processor.insights["top_products"]
            if isinstance(top_products, dict):
                for product, stats in list(top_products.items())[:3]:
                    if isinstance(stats, dict):
                        revenue = stats.get("revenue", stats.get("total_amount", 0))
                        output_lines.append(
                            f"  • {product}: {format_currency(revenue)}"
                        )

        # Restock Recommendations
        if "restock_recommendations" in processor.insights:
            output_lines.append("\n🔄 RESTOCK RECOMMENDATIONS")
            output_lines.append("-" * 40)

            for rec in processor.insights["restock_recommendations"][:3]:
                product = rec.get("product", "Unknown")
                qty = rec.get("order_quantity", 0)
                cost = rec.get("estimated_cost", 0)
                output_lines.append(
                    f"  • {product}: Order {qty:,} units ({format_currency(cost)})"
                )

        # Footer
        output_lines.append("\n" + "=" * 60)
        output_lines.append("✅ Analysis complete")
        output_lines.append("=" * 60)

        # Join all lines
        output_text = "\n".join(output_lines)
        print(output_text)

        return output_text

    @staticmethod
    def _format_key(key):
        """Format key for display."""
        return key.replace("_", " ").title()

    @staticmethod
    def _format_value(value):
        """Format value for display."""
        if isinstance(value, (int, float)):
            if abs(value) > 1000:
                return f"{value:,.0f}"
            return f"{value:.2f}"
        return str(value)
