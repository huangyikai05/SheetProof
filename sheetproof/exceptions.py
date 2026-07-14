"""User-facing exception hierarchy."""


class SheetProofError(Exception):
    """Base error for expected SheetProof failures."""


class WorkbookParseError(SheetProofError):
    """Raised when a workbook cannot be parsed safely."""


class ConfigurationError(SheetProofError):
    """Raised when a SheetProof configuration is invalid."""


class ReportGenerationError(SheetProofError):
    """Raised when a report cannot be generated."""
