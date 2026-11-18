"""Shared validation errors, independent of reader and writer construction."""


class ModelValidationError(RuntimeError):
    """A model fails validation; ``issues`` contains structured details if known."""

    def __init__(self, message, *, issues=()):
        super().__init__(message)
        self.issues = tuple(issues)


class RelationshipValidationError(ModelValidationError, ValueError):
    """Invalid relationship data/topology; retains ValueError compatibility."""
