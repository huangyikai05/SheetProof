"""User-facing exception hierarchy."""


class TabulintError(Exception):
    """Base error for expected Tabulint failures."""


class WorkbookParseError(TabulintError):
    """Raised when a workbook cannot be parsed safely."""


class ConfigurationError(TabulintError):
    """Raised when a Tabulint configuration is invalid."""


class ReportGenerationError(TabulintError):
    """Raised when a report cannot be generated."""
