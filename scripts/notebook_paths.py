"""Shared paths for workbook-scoped generated artifacts."""

from pathlib import Path


def notebook_path_for_workbook(wb):
    """Return the generated notebook path beside an xlwings workbook."""
    return Path(wb.fullname).with_suffix('.ipynb')
