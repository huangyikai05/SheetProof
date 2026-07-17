"""Bounded dependency graph construction from formula references."""

from __future__ import annotations

from tabulint.models import WorkbookSnapshot
from tabulint.parser.formulas import FormulaAnalysis, FormulaParser, qualify_cell


class DependencyGraph:
    """A small directed graph where edges point upstream -> dependent formula."""

    def __init__(self, *, max_nodes: int = 10_000) -> None:
        if max_nodes < 1:
            raise ValueError("max_nodes must be at least 1")
        self.max_nodes = max_nodes
        self._successors: dict[str, set[str]] = {}
        self._predecessors: dict[str, set[str]] = {}
        self.formula_analyses: dict[str, FormulaAnalysis] = {}
        self.unsupported_formulas: set[str] = set()
        self.truncated = False

    @classmethod
    def from_snapshot(
        cls,
        snapshot: WorkbookSnapshot,
        *,
        max_nodes: int = 10_000,
    ) -> DependencyGraph:
        return cls(max_nodes=max_nodes).build(snapshot)

    def build(self, snapshot: WorkbookSnapshot) -> DependencyGraph:
        """Populate this graph from ``snapshot`` and return ``self``."""

        self._successors.clear()
        self._predecessors.clear()
        self.formula_analyses.clear()
        self.unsupported_formulas.clear()
        self.truncated = False
        parser = FormulaParser()
        actual_sheet_names = {sheet.name.casefold(): sheet.name for sheet in snapshot.sheets}
        for sheet in sorted(snapshot.sheets, key=lambda item: item.index):
            for coordinate in sorted(sheet.cells, key=_coordinate_sort_key):
                cell = sheet.cells[coordinate]
                if cell.formula is None:
                    continue
                formula_node = qualify_cell(sheet.name, coordinate)
                if not self._add_node(formula_node):
                    self.truncated = True
                    continue
                analysis = parser.parse(cell.formula)
                self.formula_analyses[formula_node] = analysis
                if not analysis.supported:
                    self.unsupported_formulas.add(formula_node)
                stop_formula = False
                for reference in analysis.references:
                    if reference.external_workbook is not None:
                        continue
                    referenced_cells, range_truncated = reference.expand(
                        sheet.name,
                        limit=self.max_nodes,
                        qualified=True,
                    )
                    self.truncated = self.truncated or range_truncated
                    for upstream_node in referenced_cells:
                        if "!" in upstream_node:
                            upstream_sheet, upstream_coordinate = upstream_node.rsplit("!", 1)
                            actual_sheet = actual_sheet_names.get(upstream_sheet.casefold())
                            if actual_sheet is not None:
                                upstream_node = qualify_cell(actual_sheet, upstream_coordinate)
                        if not self._add_node(upstream_node):
                            self.truncated = True
                            stop_formula = True
                            break
                        self._successors[upstream_node].add(formula_node)
                        self._predecessors[formula_node].add(upstream_node)
                    if stop_formula:
                        break
        return self

    @property
    def nodes(self) -> frozenset[str]:
        return frozenset(self._successors)

    @property
    def node_count(self) -> int:
        return len(self._successors)

    @property
    def edge_count(self) -> int:
        return sum(len(values) for values in self._successors.values())

    def has_node(self, node: str) -> bool:
        return node in self._successors

    def direct_upstream(self, node: str) -> list[str]:
        return sorted(self._predecessors.get(node, ()), key=_location_sort_key)

    def direct_downstream(self, node: str) -> list[str]:
        return sorted(self._successors.get(node, ()), key=_location_sort_key)

    def successors(self, node: str) -> frozenset[str]:
        return frozenset(self._successors.get(node, ()))

    def predecessors(self, node: str) -> frozenset[str]:
        return frozenset(self._predecessors.get(node, ()))

    def _add_node(self, node: str) -> bool:
        if node in self._successors:
            return True
        if len(self._successors) >= self.max_nodes:
            return False
        self._successors[node] = set()
        self._predecessors[node] = set()
        return True


def _coordinate_sort_key(coordinate: str) -> tuple[int, int, str]:
    column = 0
    row_text = ""
    for character in coordinate.upper():
        if character.isalpha():
            column = column * 26 + (ord(character) - ord("A") + 1)
        elif character.isdigit():
            row_text += character
    return int(row_text or "0"), column, coordinate


def _location_sort_key(location: str) -> tuple[str, int, int, str]:
    if "!" not in location:
        return "", 0, 0, location
    sheet, coordinate = location.rsplit("!", 1)
    row, column, _ = _coordinate_sort_key(coordinate)
    return sheet.casefold(), row, column, coordinate


__all__ = ["DependencyGraph"]
