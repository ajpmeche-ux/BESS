"""Number and currency formatting utilities for BESS Analyzer."""

from typing import Optional


def format_currency(value: float, decimals: int = 0, prefix: str = "$") -> str:
    """Format a number as currency string.

    Args:
        value: The numeric value to format.
        decimals: Number of decimal places.
        prefix: Currency symbol prefix.

    Returns:
        Formatted currency string (e.g., "$1,234,567").
    """
    if abs(value) >= 1e9:
        return f"{prefix}{value / 1e9:,.{decimals}f}B"
    if abs(value) >= 1e6:
        return f"{prefix}{value / 1e6:,.{decimals}f}M"
    if abs(value) >= 1e3:
        return f"{prefix}{value / 1e3:,.{decimals}f}K"
    return f"{prefix}{value:,.{decimals}f}"


def format_currency_exact(value: float, decimals: int = 0, prefix: str = "$") -> str:
    """Format a number as exact currency string without abbreviation.

    Args:
        value: The numeric value to format.
        decimals: Number of decimal places.
        prefix: Currency symbol prefix.

    Returns:
        Formatted currency string (e.g., "$1,234,567").
    """
    return f"{prefix}{value:,.{decimals}f}"


def format_percent(value: float, decimals: int = 1) -> str:
    """Format a decimal as percentage string.

    Args:
        value: Decimal value (e.g., 0.07 for 7%).
        decimals: Number of decimal places.

    Returns:
        Formatted percentage string (e.g., "7.0%").
    """
    return f"{value * 100:,.{decimals}f}%"


def format_number(value: float, decimals: int = 1) -> str:
    """Format a number with comma separators.

    Args:
        value: The numeric value to format.
        decimals: Number of decimal places.

    Returns:
        Formatted number string (e.g., "1,234.5").
    """
    return f"{value:,.{decimals}f}"


def format_years(value: Optional[float]) -> str:
    """Format a value as years.

    Args:
        value: Number of years, or None if not calculable.

    Returns:
        Formatted string (e.g., "7.2 years" or "N/A").
    """
    if value is None:
        return "N/A"
    return f"{value:.1f} years"
