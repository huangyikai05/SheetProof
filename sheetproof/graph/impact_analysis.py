"""Bounded downstream impact traversal and cycle detection."""

from __future__ import annotations

from collections import deque
from typing import Any

from sheetproof.graph.dependency_graph import DependencyGraph
from sheetproof.models import DependencyImpact, WorkbookSnapshot
from sheetproof.parser.formulas import normalise_coordinate, qualify_cell, split_location


class DependencyAnalyzer:
    """Analyze changed cells against a dependency graph with hard limits."""

    def __init__(self, max_depth: int = 10, max_nodes: int = 10_000) -> None:
        if max_depth < 1:
            raise ValueError("max_depth must be at least 1")
        if max_nodes < 1:
            raise ValueError("max_nodes must be at least 1")
        self.max_depth = max_depth
        self.max_nodes = max_nodes

    def analyze(
        self,
        snapshot: WorkbookSnapshot,
        changed_cells: list[str],
        critical_cells: list[str],
        *,
        before_snapshot: WorkbookSnapshot | None = None,
    ) -> list[DependencyImpact]:
        graph = DependencyGraph.from_snapshot(snapshot, max_nodes=self.max_nodes)
        before_graph = (
            DependencyGraph.from_snapshot(before_snapshot, max_nodes=self.max_nodes)
            if before_snapshot is not None
            else None
        )
        critical = {
            resolved
            for value in critical_cells
            if (resolved := _resolve_location(snapshot, value)) is not None
        }
        results: list[DependencyImpact] = []
        seen: set[str] = set()
        for original in changed_cells:
            resolved = _resolve_location(snapshot, original)
            node = resolved or original.strip()
            before_node = (
                _resolve_location(before_snapshot, original)
                if before_snapshot is not None
                else None
            )
            if node in seen:
                continue
            seen.add(node)
            results.append(
                self._impact(
                    snapshot,
                    graph,
                    node,
                    critical,
                    valid_location=resolved is not None,
                    before_snapshot=before_snapshot,
                    before_graph=before_graph,
                    before_node=before_node,
                )
            )
        return results

    def _impact(
        self,
        snapshot: WorkbookSnapshot,
        graph: DependencyGraph,
        node: str,
        critical: set[str],
        *,
        valid_location: bool,
        before_snapshot: WorkbookSnapshot | None,
        before_graph: DependencyGraph | None,
        before_node: str | None,
    ) -> DependencyImpact:
        parents: dict[str, str | None] = {node: None}
        depths: dict[str, int] = {node: 0}
        queue: deque[str] = deque([node])
        traversal_truncated_by_depth = False
        traversal_truncated_by_nodes = False

        while queue:
            current = queue.popleft()
            depth = depths[current]
            successors = sorted(graph.successors(current), key=_location_sort_key)
            if depth >= self.max_depth:
                if any(successor not in depths for successor in successors):
                    traversal_truncated_by_depth = True
                continue
            for successor in successors:
                if successor in depths:
                    continue
                if len(depths) >= self.max_nodes:
                    traversal_truncated_by_nodes = True
                    break
                depths[successor] = depth + 1
                parents[successor] = current
                queue.append(successor)
            if traversal_truncated_by_nodes:
                break

        visited = set(depths)
        impacted = visited - {node}
        critical_impacted = sorted(visited & critical, key=_location_sort_key)
        paths = _path_examples(
            node,
            impacted,
            critical,
            parents,
            depths,
            graph,
        )
        direct_upstream_after = graph.direct_upstream(node)
        direct_upstream_before = (
            before_graph.direct_upstream(before_node)
            if before_graph is not None and before_node is not None
            else []
        )
        direct_upstream_before_set = set(direct_upstream_before)
        direct_upstream_after_set = set(direct_upstream_after)
        direct_upstream = sorted(
            direct_upstream_before_set | direct_upstream_after_set,
            key=_location_sort_key,
        )
        direct_downstream = graph.direct_downstream(node)
        involved_nodes = visited | set(direct_upstream)
        involved_sheets = {
            value.rsplit("!", 1)[0] for value in involved_nodes if "!" in value
        }
        unsupported_reachable = sorted(
            visited & graph.unsupported_formulas,
            key=_location_sort_key,
        )
        evidence: dict[str, Any] = {
            "graph_node_count": graph.node_count,
            "graph_edge_count": graph.edge_count,
            "graph_build_truncated": graph.truncated,
            "traversal_depth_limit_reached": traversal_truncated_by_depth,
            "traversal_node_limit_reached": traversal_truncated_by_nodes,
            "cell_present_in_snapshot": _cell_exists(snapshot, node),
            "valid_location": valid_location,
            "unsupported_formula_analysis": unsupported_reachable,
        }
        if before_graph is not None:
            evidence.update(
                {
                    "before_graph_node_count": before_graph.node_count,
                    "before_graph_edge_count": before_graph.edge_count,
                    "before_graph_build_truncated": before_graph.truncated,
                    "before_location_resolved": before_node is not None,
                    "before_cell_present_in_snapshot": (
                        before_snapshot is not None
                        and before_node is not None
                        and _cell_exists(before_snapshot, before_node)
                    ),
                    "direct_upstream_before": direct_upstream_before,
                    "direct_upstream_after": direct_upstream_after,
                    "direct_upstream_sources": {
                        upstream: [
                            source
                            for source, locations in (
                                ("before", direct_upstream_before_set),
                                ("after", direct_upstream_after_set),
                            )
                            if upstream in locations
                        ]
                        for upstream in direct_upstream
                    },
                    "unsupported_formula_analysis_before": sorted(
                        ({before_node} if before_node is not None else set())
                        & before_graph.unsupported_formulas,
                        key=_location_sort_key,
                    ),
                }
            )
        return DependencyImpact(
            cell=node,
            direct_upstream=direct_upstream,
            direct_downstream=direct_downstream,
            downstream_cell_count=len(impacted),
            involved_sheet_count=len(involved_sheets),
            path_examples=paths,
            critical_cells_impacted=critical_impacted,
            cycle_detected=_cycle_detected(graph, visited),
            truncated=(
                graph.truncated
                or (before_graph is not None and before_graph.truncated)
                or traversal_truncated_by_depth
                or traversal_truncated_by_nodes
            ),
            traversal_depth=max(depths.values(), default=0),
            evidence=evidence,
        )


def _resolve_location(snapshot: WorkbookSnapshot, value: str) -> str | None:
    location = value.strip()
    if not location:
        return None
    sheet_names = [sheet.name for sheet in sorted(snapshot.sheets, key=lambda item: item.index)]
    if "!" in location:
        try:
            sheet, coordinate = split_location(location)
        except ValueError:
            return None
        actual_sheet = next(
            (
                candidate.name
                for candidate in snapshot.sheets
                if candidate.name.casefold() == sheet.casefold()
            ),
            None,
        )
        return qualify_cell(actual_sheet, coordinate) if actual_sheet is not None else None
    try:
        coordinate = normalise_coordinate(location)
    except ValueError:
        return None
    containing = [
        sheet.name
        for sheet in snapshot.sheets
        if coordinate in sheet.cells
    ]
    if len(containing) == 1:
        return qualify_cell(containing[0], coordinate)
    if sheet_names:
        return qualify_cell(sheet_names[0], coordinate)
    return None


def _cell_exists(snapshot: WorkbookSnapshot, location: str) -> bool:
    if "!" not in location:
        return False
    sheet, coordinate = location.rsplit("!", 1)
    snapshot_sheet = snapshot.sheet_map().get(sheet)
    return snapshot_sheet is not None and coordinate in snapshot_sheet.cells


def _path_examples(
    start: str,
    impacted: set[str],
    critical: set[str],
    parents: dict[str, str | None],
    depths: dict[str, int],
    graph: DependencyGraph,
    *,
    limit: int = 5,
) -> list[list[str]]:
    leaves = {
        node for node in impacted if not (set(graph.successors(node)) & impacted)
    }
    candidates = sorted(
        impacted,
        key=lambda node: (
            0 if node in critical else 1,
            0 if node in leaves else 1,
            -depths[node],
            _location_sort_key(node),
        ),
    )
    result: list[list[str]] = []
    for candidate in candidates[:limit]:
        path: list[str] = []
        current: str | None = candidate
        while current is not None:
            path.append(current)
            current = parents.get(current)
        path.reverse()
        if path and path[0] == start:
            result.append(path)
    return result


def _cycle_detected(graph: DependencyGraph, nodes: set[str]) -> bool:
    """Kahn's algorithm on the bounded reachable subgraph."""

    if not nodes:
        return False
    indegree = {
        node: len(set(graph.predecessors(node)) & nodes)
        for node in nodes
    }
    queue: deque[str] = deque(
        sorted((node for node, degree in indegree.items() if degree == 0), key=_location_sort_key)
    )
    processed = 0
    while queue:
        current = queue.popleft()
        processed += 1
        for successor in sorted(graph.successors(current) & nodes, key=_location_sort_key):
            indegree[successor] -= 1
            if indegree[successor] == 0:
                queue.append(successor)
    return processed != len(nodes)


def _location_sort_key(location: str) -> tuple[str, int, int, str]:
    if "!" not in location:
        return "", 0, 0, location
    sheet, coordinate = location.rsplit("!", 1)
    column = 0
    row_text = ""
    for character in coordinate.upper():
        if character.isalpha():
            column = column * 26 + (ord(character) - ord("A") + 1)
        elif character.isdigit():
            row_text += character
    return sheet.casefold(), int(row_text or "0"), column, coordinate


__all__ = ["DependencyAnalyzer"]
