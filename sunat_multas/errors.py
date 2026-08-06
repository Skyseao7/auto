class SunatAutomationError(Exception):
    """Base exception for controlled SUNAT automation failures."""


class AuthenticationError(SunatAutomationError):
    """Raised when SUNAT rejects the company credentials."""


class NavigationError(SunatAutomationError):
    """Raised when the target SUNAT menu cannot be reached."""


class ExtractionError(SunatAutomationError):
    """Raised when data extraction fails after reaching the consultation page."""


class NoRecordsFound(SunatAutomationError):
    """Raised when the consultation succeeds but returns no records."""
