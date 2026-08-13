"""Domain models shared by the Streamlit app and processing pipeline."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import re


_INVALID_FILENAME_CHARACTERS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def sanitize_output_filename(filename: str) -> str:
    """Return a safe Excel output filename or raise ValueError."""
    cleaned_name = filename.strip()
    if not cleaned_name:
        raise ValueError("Enter an output filename.")
    if Path(cleaned_name).name != cleaned_name:
        raise ValueError("The output filename cannot include folders.")
    if _INVALID_FILENAME_CHARACTERS.search(cleaned_name):
        raise ValueError("The output filename contains unsupported characters.")
    if not cleaned_name.lower().endswith(".xlsx"):
        cleaned_name = f"{cleaned_name}.xlsx"
    if len(cleaned_name) > 120:
        raise ValueError("The output filename must be 120 characters or fewer.")
    return cleaned_name


@dataclass(frozen=True)
class RunRequest:
    """Immutable file paths and settings for one refresh."""

    master_path: Path
    source_paths: dict[str, Path]
    output_path: Path
    reporting_year: int = 2026

    def validate_paths_exist(self) -> None:
        paths = [self.master_path, *self.source_paths.values()]
        missing_paths = [str(path) for path in paths if not path.is_file()]
        if missing_paths:
            raise FileNotFoundError("Required workbook(s) not found:\n" + "\n".join(missing_paths))


@dataclass(frozen=True)
class SourceSummary:
    """Preflight metrics for one source workbook."""

    source_key: str
    display_name: str
    row_count: int
    expense_total: float


@dataclass(frozen=True)
class ValidationResult:
    """A user-displayable validation result."""

    name: str
    passed: bool
    message: str


@dataclass(frozen=True)
class ProgressEvent:
    """A progress update emitted while a refresh is running."""

    step: str
    message: str
    percent: int


@dataclass
class RunResult:
    """Final outcome of one refresh request."""

    success: bool
    output_path: Path | None = None
    validations: list[ValidationResult] = field(default_factory=list)
    source_summaries: list[SourceSummary] = field(default_factory=list)
    total_expense: float = 0.0
    error_message: str | None = None
