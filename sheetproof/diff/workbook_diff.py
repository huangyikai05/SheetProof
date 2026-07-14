"""Orchestrate deterministic structure, cell, and formula differences."""

from __future__ import annotations

from collections.abc import Sequence
from typing import Any

from sheetproof.diff.cell_diff import CellDiffer
from sheetproof.diff.formula_diff import FormulaDiffer
from sheetproof.models import (
    CellChange,
    FormulaChange,
    NamedRangeInfo,
    RiskLevel,
    SheetSnapshot,
    StructureChange,
    WorkbookSnapshot,
)


class WorkbookDiffer:
    """Compare two immutable workbook snapshots."""

    def __init__(self) -> None:
        self.cell_differ = CellDiffer()
        self.formula_differ = FormulaDiffer()

    def compare(
        self,
        before: WorkbookSnapshot,
        after: WorkbookSnapshot,
    ) -> tuple[list[StructureChange], list[CellChange], list[FormulaChange]]:
        structure_changes = self._compare_structure(before, after)
        cell_changes: list[CellChange] = []
        formula_changes: list[FormulaChange] = []
        formula_budget = self.formula_differ.new_budget()
        before_sheets = before.sheet_map()
        after_sheets = after.sheet_map()
        for sheet_name in sorted(set(before_sheets) & set(after_sheets), key=str.casefold):
            before_sheet = before_sheets[sheet_name]
            after_sheet = after_sheets[sheet_name]
            cell_changes.extend(self.cell_differ.compare(before_sheet, after_sheet))
            formula_changes.extend(
                self.formula_differ.compare(
                    before_sheet,
                    after_sheet,
                    budget=formula_budget,
                )
            )
        for sheet_name in sorted(set(after_sheets) - set(before_sheets), key=str.casefold):
            after_sheet = after_sheets[sheet_name]
            empty_before = after_sheet.model_copy(update={"cells": {}})
            cell_changes.extend(self.cell_differ.compare(empty_before, after_sheet))
            formula_changes.extend(
                self.formula_differ.compare(
                    empty_before,
                    after_sheet,
                    budget=formula_budget,
                )
            )
        for sheet_name in sorted(set(before_sheets) - set(after_sheets), key=str.casefold):
            before_sheet = before_sheets[sheet_name]
            empty_after = before_sheet.model_copy(update={"cells": {}})
            cell_changes.extend(self.cell_differ.compare(before_sheet, empty_after))
            formula_changes.extend(
                self.formula_differ.compare(
                    before_sheet,
                    empty_after,
                    budget=formula_budget,
                )
            )
        return structure_changes, cell_changes, formula_changes

    def _compare_structure(
        self,
        before: WorkbookSnapshot,
        after: WorkbookSnapshot,
    ) -> list[StructureChange]:
        changes: list[StructureChange] = []
        before_sheets = before.sheet_map()
        after_sheets = after.sheet_map()
        before_names = [sheet.name for sheet in before.sheets]
        after_names = [sheet.name for sheet in after.sheets]

        for name in sorted(set(after_sheets) - set(before_sheets), key=str.casefold):
            sheet = after_sheets[name]
            hidden = sheet.state != "visible"
            changes.append(
                _change(
                    "hidden_sheet_added" if hidden else "sheet_added",
                    RiskLevel.HIGH if hidden else RiskLevel.MEDIUM,
                    f"sheet:{name}",
                    None,
                    _sheet_summary(sheet),
                    (
                        f"Hidden worksheet '{name}' was added"
                        if hidden
                        else f"Worksheet '{name}' was added"
                    ),
                    {"sheet": name, "state": sheet.state, "index": sheet.index},
                )
            )
        for name in sorted(set(before_sheets) - set(after_sheets), key=str.casefold):
            sheet = before_sheets[name]
            hidden = sheet.state != "visible"
            changes.append(
                _change(
                    "hidden_sheet_removed" if hidden else "sheet_deleted",
                    RiskLevel.MEDIUM if hidden else RiskLevel.HIGH,
                    f"sheet:{name}",
                    _sheet_summary(sheet),
                    None,
                    (
                        f"Hidden worksheet '{name}' was removed"
                        if hidden
                        else f"Worksheet '{name}' was deleted"
                    ),
                    {"sheet": name, "state": sheet.state, "index": sheet.index},
                )
            )

        common_before_order = [name for name in before_names if name in after_sheets]
        common_after_order = [name for name in after_names if name in before_sheets]
        if common_before_order != common_after_order:
            changes.append(
                _change(
                    "sheet_order_changed",
                    RiskLevel.LOW,
                    "workbook:sheets",
                    common_before_order,
                    common_after_order,
                    "Worksheet order changed",
                    {
                        "before_order": common_before_order,
                        "after_order": common_after_order,
                    },
                )
            )

        for name in sorted(set(before_sheets) & set(after_sheets), key=str.casefold):
            changes.extend(_compare_sheet_structure(before_sheets[name], after_sheets[name]))

        changes.extend(_compare_named_ranges(before.named_ranges, after.named_ranges))
        before_links = set(before.external_links)
        after_links = set(after.external_links)
        for link in sorted(after_links - before_links):
            changes.append(
                _change(
                    "external_link_added",
                    RiskLevel.HIGH,
                    "workbook:external_links",
                    None,
                    link,
                    f"External workbook link was added: {link}",
                    {"link": link},
                )
            )
        for link in sorted(before_links - after_links):
            changes.append(
                _change(
                    "external_link_removed",
                    RiskLevel.LOW,
                    "workbook:external_links",
                    link,
                    None,
                    f"External workbook link was removed: {link}",
                    {"link": link},
                )
            )
        if before.has_vba != after.has_vba:
            added = after.has_vba
            changes.append(
                _change(
                    "macro_added" if added else "macro_removed",
                    RiskLevel.CRITICAL if added else RiskLevel.MEDIUM,
                    "workbook:vba",
                    before.has_vba,
                    after.has_vba,
                    "VBA macro content was added" if added else "VBA macro content was removed",
                    {"before_has_vba": before.has_vba, "after_has_vba": after.has_vba},
                )
            )
        return changes


def _compare_sheet_structure(
    before: SheetSnapshot,
    after: SheetSnapshot,
) -> list[StructureChange]:
    changes: list[StructureChange] = []
    location = f"sheet:{before.name}"
    if before.state != after.state:
        newly_hidden = before.state == "visible" and after.state != "visible"
        changes.append(
            _change(
                "sheet_visibility_changed",
                RiskLevel.HIGH if newly_hidden else RiskLevel.MEDIUM,
                location,
                before.state,
                after.state,
                (
                    f"Worksheet '{before.name}' visibility changed "
                    f"from {before.state} to {after.state}"
                ),
                {
                    "sheet": before.name,
                    "before_state": before.state,
                    "after_state": after.state,
                    "newly_hidden": newly_hidden,
                },
            )
        )
    changes.extend(
        _sequence_change(
            "hidden_rows_changed",
            RiskLevel.MEDIUM,
            f"{location}:rows",
            before.hidden_rows,
            after.hidden_rows,
            f"Hidden row state changed on worksheet '{before.name}'",
            "rows",
        )
    )
    changes.extend(
        _sequence_change(
            "hidden_columns_changed",
            RiskLevel.MEDIUM,
            f"{location}:columns",
            before.hidden_columns,
            after.hidden_columns,
            f"Hidden column state changed on worksheet '{before.name}'",
            "columns",
        )
    )
    changes.extend(
        _sequence_change(
            "merged_cells_changed",
            RiskLevel.MEDIUM,
            f"{location}:merged_cells",
            before.merged_cells,
            after.merged_cells,
            f"Merged cell ranges changed on worksheet '{before.name}'",
            "ranges",
        )
    )

    before_validations = [item.model_dump(mode="json") for item in before.data_validations]
    after_validations = [item.model_dump(mode="json") for item in after.data_validations]
    if before_validations != after_validations:
        changes.append(
            _change(
                "data_validations_changed",
                RiskLevel.MEDIUM,
                f"{location}:data_validations",
                before_validations,
                after_validations,
                f"Data validation rules changed on worksheet '{before.name}'",
                {
                    "before_count": len(before_validations),
                    "after_count": len(after_validations),
                },
            )
        )
    if before.freeze_panes != after.freeze_panes:
        changes.append(
            _change(
                "freeze_panes_changed",
                RiskLevel.LOW,
                f"{location}:freeze_panes",
                before.freeze_panes,
                after.freeze_panes,
                f"Freeze panes changed on worksheet '{before.name}'",
                {"before": before.freeze_panes, "after": after.freeze_panes},
            )
        )
    before_tables = [item.model_dump(mode="json") for item in before.tables]
    after_tables = [item.model_dump(mode="json") for item in after.tables]
    if before_tables != after_tables:
        changes.append(
            _change(
                "tables_changed",
                RiskLevel.MEDIUM,
                f"{location}:tables",
                before_tables,
                after_tables,
                f"Table definitions changed on worksheet '{before.name}'",
                {"before_count": len(before_tables), "after_count": len(after_tables)},
            )
        )
    return changes


def _compare_named_ranges(
    before: Sequence[NamedRangeInfo],
    after: Sequence[NamedRangeInfo],
) -> list[StructureChange]:
    before_map = {(item.name, item.local_sheet_id): item for item in before}
    after_map = {(item.name, item.local_sheet_id): item for item in after}
    changes: list[StructureChange] = []
    keys = sorted(
        set(before_map) | set(after_map),
        key=lambda key: (
            key[0].casefold(),
            -1 if key[1] is None else key[1],
        ),
    )
    for key in keys:
        before_item = before_map.get(key)
        after_item = after_map.get(key)
        location = f"named_range:{key[0]}"
        if key[1] is not None:
            location = f"{location}:sheet_index={key[1]}"
        if before_item is None and after_item is not None:
            changes.append(
                _change(
                    "named_range_added",
                    RiskLevel.MEDIUM,
                    location,
                    None,
                    after_item.model_dump(mode="json"),
                    f"Named range '{key[0]}' was added",
                    {"name": key[0], "value": after_item.value},
                )
            )
        elif before_item is not None and after_item is None:
            changes.append(
                _change(
                    "named_range_removed",
                    RiskLevel.HIGH,
                    location,
                    before_item.model_dump(mode="json"),
                    None,
                    f"Named range '{key[0]}' was removed",
                    {"name": key[0], "value": before_item.value},
                )
            )
        elif before_item is not None and after_item is not None and before_item != after_item:
            changes.append(
                _change(
                    "named_range_changed",
                    RiskLevel.HIGH,
                    location,
                    before_item.model_dump(mode="json"),
                    after_item.model_dump(mode="json"),
                    f"Named range '{key[0]}' changed",
                    {
                        "name": key[0],
                        "before_value": before_item.value,
                        "after_value": after_item.value,
                    },
                )
            )
    return changes


def _sequence_change(
    change_type: str,
    risk_level: RiskLevel,
    location: str,
    before: Sequence[str] | Sequence[int],
    after: Sequence[str] | Sequence[int],
    description: str,
    evidence_name: str,
) -> list[StructureChange]:
    if list(before) == list(after):
        return []
    before_set = set(before)
    after_set = set(after)
    return [
        _change(
            change_type,
            risk_level,
            location,
            list(before),
            list(after),
            description,
            {
                f"added_{evidence_name}": sorted(after_set - before_set),
                f"removed_{evidence_name}": sorted(before_set - after_set),
            },
        )
    ]


def _sheet_summary(sheet: SheetSnapshot) -> dict[str, Any]:
    return {
        "name": sheet.name,
        "index": sheet.index,
        "state": sheet.state,
        "cell_count": len(sheet.cells),
    }


def _change(
    change_type: str,
    risk_level: RiskLevel,
    location: str,
    before: Any,
    after: Any,
    description: str,
    evidence: dict[str, Any],
) -> StructureChange:
    return StructureChange(
        change_type=change_type,
        risk_level=risk_level,
        location=location,
        before=before,
        after=after,
        description=description,
        evidence=evidence,
    )


__all__ = ["WorkbookDiffer"]
