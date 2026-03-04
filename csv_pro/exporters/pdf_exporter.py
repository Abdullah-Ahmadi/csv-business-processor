from datetime import datetime
from fpdf import FPDF
from .base_exporter import BaseExporter
from utils.formatters import format_currency, format_percentage
from utils.file_utils import generate_filename
from utils.logger import get_logger


class ReportPDF(FPDF):
    """
    Custom FPDF subclass to ensure footer
    and consistent page framing.
    """

    def footer(self):
        # Skip footer on cover page
        if self.page_no() == 1:
            return

        self.set_y(-15)
        self.set_font("Helvetica", "", 8)
        self.set_draw_color(200, 200, 200)

        # Footer separator
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(2)

        # Page numbering starts from content page
        page_number = self.page_no() - 1

        self.cell(
            0,
            10,
            f"Page {page_number} - Confidential Business Report",
            0,
            0,
            "C",
        )


class PDFExporter(BaseExporter):
    """
    Export processor insights to PDF format using FPDF.
    """

    PAGE_WIDTH = 190
    LINE_HEIGHT = 6
    SECTION_GAP = 10

    def __init__(self):
        self.logger = get_logger(self.__class__.__name__)
        self.pdf = None

    def export(self, processor, output_path=None):

        self._validate_processor(processor)
        output_path = self._resolve_output_path(processor, output_path)

        self._create_document()
        self._add_cover_page(processor)
        self._compose_document(processor)

        self._save(output_path)

        return output_path

    def _create_document(self):
        """Create PDF document."""
        self.pdf = ReportPDF(format="A4")
        self.pdf.set_auto_page_break(auto=True, margin=20)
        self.pdf.set_margins(10, 15, 10)
        self.pdf.set_font("Helvetica", "", 10)

    def _draw_page_frame(self):
        """Draw a frame around the page for visual appeal."""
        self.pdf.set_draw_color(200, 200, 200)
        self.pdf.rect(5, 5, 200, 287)

    def _add_cover_page(self, processor):
        """Add a cover page to the PDF."""
        self.pdf.add_page()
        self._draw_page_frame()

        # Title
        self.pdf.set_font("Helvetica", "B", 24)
        self.pdf.ln(60)

        title = (
            f"{processor.__class__.__name__.replace('Processor', '')} "
            f"Analysis Report"
        )

        self.pdf.cell(0, 12, title, ln=True, align="C")

        self.pdf.ln(10)

        # Subtitle
        self.pdf.set_font("Helvetica", "", 14)
        self.pdf.cell(0, 10, "Confidential Business Report", ln=True, align="C")

        self.pdf.ln(20)

        # Date
        self.pdf.set_font("Helvetica", "", 12)
        self.pdf.cell(
            0,
            10,
            f"Generated on: {datetime.now().strftime('%Y-%m-%d')}",
            ln=True,
            align="C",
        )

    def _resolve_output_path(self, processor, output_path):
        """Resolve the output file path."""
        if output_path is None:
            output_path = generate_filename(processor, ".pdf")

        if not output_path.endswith(".pdf"):
            output_path += ".pdf"

        return output_path

    def _compose_document(self, processor):
        """Compose the main content of the PDF."""
        self.pdf.add_page()
        self._draw_page_frame()

        self._add_title(processor)
        self._add_metadata(processor)
        self._add_summary(processor)
        self._add_key_insights(processor)
        self._add_alerts(processor)
        self._add_top_products(processor)
        self._add_restock_recommendations(processor)

    def _add_title(self, processor):
        """Add title section."""
        self.pdf.set_font("Helvetica", "B", 16)

        title = (
            f"{processor.__class__.__name__.replace('Processor', '')} "
            f"Analysis Report"
        )

        self.pdf.set_fill_color(30, 60, 120)
        self.pdf.set_text_color(255, 255, 255)
        self.pdf.cell(self.PAGE_WIDTH, 12, title, ln=True, align="C", fill=True)

        self.pdf.set_text_color(0, 0, 0)
        self.pdf.ln(self.SECTION_GAP)

    def _add_metadata(self, processor):
        """Add metadata section."""
        self.pdf.set_font("Helvetica", "", 10)

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        self.pdf.set_fill_color(240, 240, 240)
        self.pdf.multi_cell(
            self.PAGE_WIDTH,
            self.LINE_HEIGHT,
            f"Generated: {timestamp}\n" f"Processor: {processor.__class__.__name__}",
            fill=True,
        )

        self.pdf.ln(self.SECTION_GAP)

    def _add_summary(self, processor):
        """Add executive summary section."""
        self._section_header("EXECUTIVE SUMMARY")

        for key, value in processor.get_processing_summary().items():
            label = key.replace("_", " ").title()
            self.pdf.cell(
                self.PAGE_WIDTH,
                self.LINE_HEIGHT,
                f"{label}: {f"{value[:100]}..." if isinstance(value, str) and len(value) > 50 else value}",
                ln=True,
            )

        self.pdf.ln(self.SECTION_GAP)

    def _add_key_insights(self, processor):
        """Add key insights section - NO BULLETS."""
        self._section_header("KEY INSIGHTS")

        count = 0

        for key, value in processor.insights.items():
            if isinstance(value, (int, float)) and not isinstance(value, bool):

                label = key.replace("_", " ").title()
                formatted_value = self._format_value(key, value)

                # Using hyphen instead of bullet
                self.pdf.cell(
                    self.PAGE_WIDTH,
                    self.LINE_HEIGHT,
                    f"- {label}: {formatted_value}",
                    ln=True,
                )

                count += 1
                if count >= 15:
                    self.pdf.cell(
                        self.PAGE_WIDTH,
                        self.LINE_HEIGHT,
                        "... additional insights available in Excel/JSON export",
                        ln=True,
                    )
                    break

        self.pdf.ln(self.SECTION_GAP)

    def _add_alerts(self, processor):
        """Add alerts section if any exist - NO BULLETS."""
        if not processor.alerts:
            return

        self._section_header("ALERTS & RECOMMENDATIONS")

        self.pdf.set_font("Helvetica", "", 10)

        for alert in processor.alerts[:8]:
            # Using hyphen instead of bullet
            self.pdf.cell(
                self.PAGE_WIDTH,
                self.LINE_HEIGHT,
                f"- {alert}",
                ln=True,
            )

        self.pdf.ln(self.SECTION_GAP)

    def _add_top_products(self, processor):
        """Add top products section if available."""
        top_products = processor.insights.get("top_products")
        if not isinstance(top_products, dict):
            return

        self._section_header("TOP PERFORMING PRODUCTS")

        for i, (product, stats) in enumerate(list(top_products.items())[:5], start=1):
            if isinstance(stats, dict):
                revenue = stats.get(
                    "revenue",
                    stats.get("total_amount", 0),
                )

                self.pdf.cell(
                    self.PAGE_WIDTH,
                    self.LINE_HEIGHT,
                    f"{i}. {product}: {format_currency(revenue)}",
                    ln=True,
                )

    def _add_restock_recommendations(self, processor):
        """Add restock recommendations section if available."""
        recs = processor.insights.get("restock_recommendations")
        if not recs:
            return

        self.pdf.add_page()
        self._draw_page_frame()

        self._section_header("RESTOCK RECOMMENDATIONS")

        recs = recs[:10]

        for rec in recs:
            product = rec.get("product", "Unknown")
            qty = rec.get("order_quantity", 0)
            cost = rec.get("estimated_cost", 0)

            # Using hyphen instead of bullet
            self.pdf.cell(
                self.PAGE_WIDTH,
                self.LINE_HEIGHT,
                f"- {product}: Order {qty} units ({format_currency(cost)})",
                ln=True,
            )

        total_cost = sum(rec.get("estimated_cost", 0) for rec in recs)

        self.pdf.ln(5)
        self.pdf.set_font("Helvetica", "B", 10)

        self.pdf.cell(
            self.PAGE_WIDTH,
            self.LINE_HEIGHT,
            f"Total Estimated Restock Cost: {format_currency(total_cost)}",
            ln=True,
        )

    def _section_header(self, title):
        """Add a section header with underline."""
        self.pdf.set_font("Helvetica", "B", 12)

        self.pdf.set_draw_color(180, 180, 180)
        self.pdf.line(10, self.pdf.get_y(), 200, self.pdf.get_y())
        self.pdf.ln(4)

        self.pdf.set_fill_color(230, 230, 230)
        self.pdf.cell(self.PAGE_WIDTH, 8, title, ln=True, fill=True)

        self.pdf.set_font("Helvetica", "", 10)
        self.pdf.ln(5)

    def _format_value(self, key, value):
        """Format values based on key name."""
        key_lower = key.lower()

        if any(
            term in key_lower
            for term in [
                "revenue",
                "amount",
                "cost",
                "price",
                "spent",
                "value",
                "total",
            ]
        ):
            return format_currency(value)

        if any(
            term in key_lower for term in ["percentage", "rate", "growth", "margin"]
        ):
            return format_percentage(value)

        return f"{value:,}"

    def _save(self, output_path):
        """Save the PDF to file."""
        self.pdf.output(output_path)
        print(f"✅ PDF report saved to: {output_path}")
