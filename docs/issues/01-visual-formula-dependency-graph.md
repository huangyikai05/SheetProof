# Issue draft: Visual formula dependency graph

Suggested labels: `enhancement`, `help wanted`

## Background

Tabulint already builds a bounded dependency graph and reports deterministic impact evidence.
That evidence is currently easier to inspect as structured data than as a relationship map. A
small visual graph would help reviewers see which changed formula feeds each impacted cell without
changing the review verdict.

## Goal

Add an offline visualization generated only from the existing typed dependency and impact facts.
The same input must produce the same node order, edges, labels, and truncation notice.

## Suggested implementation

- Define a presentation model derived from `ReviewResult`; do not add graph decisions to a report
  template or interface adapter.
- Render a compact SVG or an offline HTML section with changed cells, directly impacted cells, and
  transitively impacted cells visually distinguished.
- Reuse existing traversal limits. Sort nodes and edges before rendering, and state when a graph is
  truncated.
- Escape workbook-controlled sheet names, formulas, and cell values in every output context.
- Add a synthetic example that contains both direct and transitive dependencies.

## Acceptance criteria

- [ ] A synthetic formula change produces the expected nodes and directed edges.
- [ ] Repeated runs generate byte-stable or structurally identical graph output.
- [ ] Existing dependency limits are enforced and truncation is visible to the reviewer.
- [ ] Hostile sheet names and formula text are safely escaped in offline HTML/SVG.
- [ ] Tests cover an empty graph, a normal graph, a cycle, and a truncated graph.
- [ ] Documentation explains the legend and how the graph relates to existing evidence.

## Non-goals

- Calculating formulas or inferring values Excel would produce.
- Using graph layout, AI output, or visual prominence to determine a verdict or risk score.
- A hosted graph service, workbook upload, or an unbounded whole-workbook visualization.
