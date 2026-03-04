"""
Data Formatting utilities
"""


def format_currency(amount):
    return f"${amount:,.2f}"


def format_percentage(value):
    return f"{value:.1f}%"


def format_date(date):
    return date


def truncate_text(text, max_length=50):
    if len(text) > max_length:
        return text[: max_length - 3] + "..."
    else:
        return text
