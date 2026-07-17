"""Typed contracts shared by every Tabulint interface."""

from __future__ import annotations

import re
from datetime import UTC, datetime
from enum import StrEnum
from pathlib import Path
from typing import Annotated, Any, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

ScalarValue = str | int | float | bool | None

DEFAULT_RISK_WEIGHT_VALUES: dict[str, int] = {
    "macro_added": 40,
    "external_link_added": 35,
    "formula_overwritten": 30,
    "formula_deleted": 30,
    "hidden_sheet_added": 20,
    "hidden_rows_changed": 10,
    "hidden_columns_changed": 10,
    "formula_range_reduced": 20,
    "formula_changed": 15,
    "formula_added": 8,
    "many_cells_changed": 10,
    "text_changed": 1,
    "style_changed": 1,
    "value_changed": 2,
}
CANONICAL_RISK_WEIGHT_KEYS = frozenset(DEFAULT_RISK_WEIGHT_VALUES)

RISK_WEIGHT_ALIASES: dict[str, str] = {
    "vba_added": "macro_added",
    "macro_status_added": "macro_added",
    "external_links_added": "external_link_added",
    "new_external_link": "external_link_added",
    "new_hidden_sheet": "hidden_sheet_added",
    "row_visibility_changed": "hidden_rows_changed",
    "hidden_row_changed": "hidden_rows_changed",
    "column_visibility_changed": "hidden_columns_changed",
    "hidden_column_changed": "hidden_columns_changed",
    "formula_overwrite": "formula_overwritten",
    "formula_to_value": "formula_overwritten",
    "formula_to_fixed_value": "formula_overwritten",
    "formula_to_number": "formula_overwritten",
    "formula_to_text": "formula_overwritten",
    "formula_replaced_with_value": "formula_overwritten",
    "formula_to_blank": "formula_deleted",
    "formula_removed": "formula_deleted",
    "formula_reference_changed": "formula_changed",
    "formula_references_changed": "formula_changed",
    "fixed_value_to_formula": "formula_added",
    "blank_to_formula": "formula_added",
    "bulk_cell_changes": "many_cells_changed",
    "large_cell_change": "many_cells_changed",
    "format_changed": "style_changed",
    "numeric_value_changed": "value_changed",
    "date_changed": "value_changed",
    "error_value_changed": "value_changed",
    "data_type_changed": "value_changed",
    "cached_value_changed": "value_changed",
    "cell_cleared": "value_changed",
    "cell_populated": "value_changed",
    "blank_to_number": "value_changed",
    "blank_to_text": "text_changed",
}

_RISK_WEIGHT_NAME_RE = re.compile(r"[^a-z0-9]+")


def canonical_risk_weight_name(value: str) -> str:
    """Return the risk scorer's stable name for a configured weight key."""

    normalized = _RISK_WEIGHT_NAME_RE.sub("_", value.strip().lower()).strip("_")
    return RISK_WEIGHT_ALIASES.get(normalized, normalized)


class StrictModel(BaseModel):
    """Base model that rejects misspelled input fields."""

    model_config = ConfigDict(extra="forbid")


class RiskLevel(StrEnum):
    LOW = "LOW"
    MEDIUM = "MEDIUM"
    HIGH = "HIGH"
    CRITICAL = "CRITICAL"


class Severity(StrEnum):
    LOW = "low"
    MEDIUM = "medium"
    HIGH = "high"
    CRITICAL = "critical"


class RuleStatus(StrEnum):
    PASSED = "PASSED"
    FAILED = "FAILED"
    WARNING = "WARNING"
    SKIPPED = "SKIPPED"
    ERROR = "ERROR"


class RuleFailureStatus(StrEnum):
    FAILED = "FAILED"
    WARNING = "WARNING"


class CellKind(StrEnum):
    BLANK = "blank"
    FORMULA = "formula"
    NUMBER = "number"
    TEXT = "text"
    BOOLEAN = "boolean"
    DATE = "date"
    ERROR = "error"
    OTHER = "other"


class FormulaCalculationStatus(StrEnum):
    CACHED = "cached"
    UNCALCULATED = "uncalculated"


class RuleType(StrEnum):
    FORMULA_REQUIRED = "formula_required"
    ALLOWED_CHANGE_RANGE = "allowed_change_range"
    NO_EXTERNAL_LINKS = "no_external_links"
    NO_NEW_HIDDEN_SHEETS = "no_new_hidden_sheets"
    NO_MACRO_ADDED = "no_macro_added"
    NUMERIC_RANGE = "numeric_range"
    REQUIRED_SHEET = "required_sheet"
    FORBIDDEN_SHEET = "forbidden_sheet"
    MAX_CHANGED_CELLS = "max_changed_cells"


class FileInfo(StrictModel):
    path: str
    name: str
    suffix: str
    size_bytes: int
    sha256: str
    modified_at: datetime
    has_vba: bool

    @classmethod
    def from_path(cls, path: Path, *, sha256: str, has_vba: bool) -> FileInfo:
        stat = path.stat()
        return cls(
            path=str(path),
            name=path.name,
            suffix=path.suffix.lower(),
            size_bytes=stat.st_size,
            sha256=sha256,
            modified_at=datetime.fromtimestamp(stat.st_mtime, tz=UTC),
            has_vba=has_vba,
        )


class StyleSummary(StrictModel):
    style_id: int
    number_format: str
    font: str
    fill: str
    border: str
    alignment: str
    protection: str


class CellSnapshot(StrictModel):
    coordinate: str
    value: ScalarValue
    formula: str | None = None
    formula_attributes: dict[str, ScalarValue] = Field(default_factory=dict)
    cached_value: ScalarValue = None
    calculation_status: FormulaCalculationStatus | None = None
    kind: CellKind
    data_type: str
    is_date: bool = False
    style: StyleSummary


class DataValidationInfo(StrictModel):
    ranges: str
    validation_type: str | None
    operator: str | None
    formula1: str | None
    formula2: str | None
    allow_blank: bool
    show_error_message: bool


class TableInfo(StrictModel):
    name: str
    display_name: str
    ref: str
    totals_row_shown: bool | None


class NamedRangeInfo(StrictModel):
    name: str
    value: str
    kind: str | None
    local_sheet_id: int | None
    hidden: bool | None


class SheetSnapshot(StrictModel):
    name: str
    index: int
    state: Literal["visible", "hidden", "veryHidden"] | str
    cells: dict[str, CellSnapshot]
    hidden_rows: list[int]
    hidden_columns: list[str]
    merged_cells: list[str]
    data_validations: list[DataValidationInfo]
    freeze_panes: str | None
    tables: list[TableInfo]


class WorkbookSnapshot(StrictModel):
    file: FileInfo
    sheets: list[SheetSnapshot]
    named_ranges: list[NamedRangeInfo]
    external_links: list[str]
    has_vba: bool
    parsed_cell_count: int

    def sheet_map(self) -> dict[str, SheetSnapshot]:
        return {sheet.name: sheet for sheet in self.sheets}


class StructureChange(StrictModel):
    change_type: str
    risk_level: RiskLevel
    location: str
    before: Any = None
    after: Any = None
    description: str
    evidence: dict[str, Any] = Field(default_factory=dict)


class CellChange(StrictModel):
    change_type: str
    risk_level: RiskLevel
    sheet: str
    coordinate: str
    location: str
    before: CellSnapshot | None
    after: CellSnapshot | None
    description: str
    evidence: dict[str, Any] = Field(default_factory=dict)


class FormulaChange(StrictModel):
    change_type: str
    risk_level: RiskLevel
    sheet: str
    coordinate: str
    location: str
    before_formula: str | None
    after_formula: str | None
    replacement_value: ScalarValue = None
    replacement_kind: CellKind | None = None
    manual_review_recommendation: str | None = None
    description: str
    supported_analysis: bool
    high_impact: bool
    references_before: list[str] = Field(default_factory=list)
    references_after: list[str] = Field(default_factory=list)
    excluded_references: list[str] = Field(default_factory=list)
    neighboring_formula_pattern: list[str] = Field(default_factory=list)
    evidence: dict[str, Any] = Field(default_factory=dict)


class DependencyImpact(StrictModel):
    cell: str
    direct_upstream: list[str]
    direct_downstream: list[str]
    downstream_cell_count: int
    involved_sheet_count: int
    path_examples: list[list[str]]
    critical_cells_impacted: list[str]
    cycle_detected: bool
    truncated: bool
    traversal_depth: int
    evidence: dict[str, Any] = Field(default_factory=dict)


class RuleResult(StrictModel):
    name: str
    rule_type: str
    status: RuleStatus
    severity: Severity
    reason: str
    location: str | None = None
    evidence: dict[str, Any] = Field(default_factory=dict)


class RiskFactor(StrictModel):
    risk_type: str
    location: str
    points: int
    description: str
    evidence: dict[str, Any] = Field(default_factory=dict)


class ReviewSummary(StrictModel):
    risk_score: Annotated[int, Field(ge=0, le=100)]
    risk_level: RiskLevel
    changed_cells: int
    changed_formulas: int
    formula_overwrites: int
    added_hidden_sheets: int
    added_external_links: int
    macro_status_changed: bool
    rules_passed: int
    rules_failed: int
    rules_warnings: int
    rules_skipped: int
    rules_errors: int


class ReviewResult(StrictModel):
    tool_version: str
    reviewed_at: datetime
    before_file: FileInfo
    after_file: FileInfo
    summary: ReviewSummary
    structure_changes: list[StructureChange]
    cell_changes: list[CellChange]
    formula_changes: list[FormulaChange]
    dependency_impacts: list[DependencyImpact]
    rule_results: list[RuleResult]
    risk_factors: list[RiskFactor]
    limitations: list[str]
    errors: list[str]


class RuleSpec(StrictModel):
    name: str
    type: RuleType
    severity: Severity = Severity.HIGH
    failure_status: RuleFailureStatus = RuleFailureStatus.FAILED
    range: str | None = None
    ranges: list[str] = Field(default_factory=list)
    target: str | None = None
    sheet: str | None = None
    min: Annotated[float, Field(strict=True, allow_inf_nan=False)] | None = None
    max: Annotated[float, Field(strict=True, allow_inf_nan=False)] | None = None

    @model_validator(mode="after")
    def validate_required_fields(self) -> RuleSpec:
        if self.type is RuleType.FORMULA_REQUIRED and not self.range:
            raise ValueError("formula_required requires 'range'")
        if self.type is RuleType.ALLOWED_CHANGE_RANGE and not self.ranges:
            raise ValueError("allowed_change_range requires non-empty 'ranges'")
        if self.type is RuleType.NUMERIC_RANGE:
            if not self.target:
                raise ValueError("numeric_range requires 'target'")
            if self.min is None and self.max is None:
                raise ValueError("numeric_range requires 'min' and/or 'max'")
        if self.type in {RuleType.REQUIRED_SHEET, RuleType.FORBIDDEN_SHEET} and not self.sheet:
            raise ValueError(f"{self.type.value} requires 'sheet'")
        if self.type is RuleType.MAX_CHANGED_CELLS and (
            self.max is None or self.max < 0 or not float(self.max).is_integer()
        ):
            raise ValueError("max_changed_cells requires a non-negative integer 'max'")
        return self


class TabulintConfig(StrictModel):
    rules: list[RuleSpec] = Field(default_factory=list)
    risk_weights: dict[
        str,
        Annotated[int, Field(strict=True, ge=0, le=100)],
    ] = Field(
        default_factory=dict
    )
    block_risk_level: RiskLevel = RiskLevel.HIGH
    critical_cells: list[str] = Field(default_factory=list)
    max_dependency_depth: Annotated[int, Field(strict=True, ge=1, le=100)] = 10
    max_dependency_nodes: Annotated[
        int,
        Field(strict=True, ge=10, le=100_000),
    ] = 10_000

    @model_validator(mode="after")
    def validate_risk_weight_names(self) -> TabulintConfig:
        unknown = sorted(
            name
            for name in self.risk_weights
            if canonical_risk_weight_name(name) not in CANONICAL_RISK_WEIGHT_KEYS
        )
        if unknown:
            names = ", ".join(repr(name) for name in unknown)
            raise ValueError(f"unknown risk weight name(s): {names}")
        return self
