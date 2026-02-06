"""Input validation functions for BESS Analyzer.

Each validator returns a tuple of (is_valid: bool, message: str).
Messages describe errors or warnings for user display.
"""

from typing import List, Tuple

from src.models.project import Project


def validate_capacity(capacity_mw: float) -> Tuple[bool, str]:
    """Validate power capacity in MW.

    Args:
        capacity_mw: Project nameplate capacity.

    Returns:
        (is_valid, message) tuple.
    """
    if capacity_mw <= 0:
        return False, "Capacity must be greater than 0 MW."
    if capacity_mw > 1000:
        return True, f"Warning: {capacity_mw} MW is unusually large. Verify this is correct."
    return True, ""


def validate_duration(duration_hours: float) -> Tuple[bool, str]:
    """Validate storage duration in hours.

    Args:
        duration_hours: Storage duration.

    Returns:
        (is_valid, message) tuple.
    """
    if duration_hours <= 0:
        return False, "Duration must be greater than 0 hours."
    if duration_hours < 1:
        return True, "Warning: Duration < 1 hour is unusual for utility-scale BESS."
    if duration_hours > 24:
        return True, "Warning: Duration > 24 hours is unusual. Verify this is correct."
    return True, ""


def validate_efficiency(efficiency: float) -> Tuple[bool, str]:
    """Validate round-trip efficiency.

    Args:
        efficiency: RTE as decimal (e.g., 0.85 for 85%).

    Returns:
        (is_valid, message) tuple.
    """
    if efficiency < 0.70:
        return False, "Round-trip efficiency must be at least 70%."
    if efficiency > 0.95:
        return False, "Round-trip efficiency cannot exceed 95%."
    return True, ""


def validate_discount_rate(rate: float) -> Tuple[bool, str]:
    """Validate discount rate.

    Args:
        rate: Discount rate as decimal (e.g., 0.07 for 7%).

    Returns:
        (is_valid, message) tuple.
    """
    if rate < 0.01:
        return False, "Discount rate must be at least 1%."
    if rate > 0.20:
        return False, "Discount rate cannot exceed 20%."
    return True, ""


def validate_capex(capex_per_kwh: float) -> Tuple[bool, str]:
    """Validate capital cost per kWh."""
    if capex_per_kwh <= 0:
        return False, "CapEx must be greater than $0/kWh."
    if capex_per_kwh > 500:
        return True, "Warning: CapEx > $500/kWh is above current market range."
    return True, ""


def validate_project(project: Project) -> Tuple[bool, List[str]]:
    """Run all validations on a complete project.

    Args:
        project: Project to validate.

    Returns:
        (is_valid, messages) where messages includes all errors and warnings.
    """
    messages = []
    is_valid = True

    checks = [
        validate_capacity(project.basics.capacity_mw),
        validate_duration(project.basics.duration_hours),
        validate_efficiency(project.technology.round_trip_efficiency),
        validate_discount_rate(project.basics.discount_rate),
        validate_capex(project.costs.capex_per_kwh),
    ]

    for valid, msg in checks:
        if not valid:
            is_valid = False
        if msg:
            messages.append(msg)

    if not project.basics.name.strip():
        messages.append("Warning: Project name is empty.")

    if len(project.benefits) == 0:
        messages.append("Warning: No benefit streams defined. NPV will be negative.")

    return is_valid, messages
