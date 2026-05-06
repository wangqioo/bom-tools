---
name: pstx-cli-analysis
description: Use the dehdl_review PSTX Trae Bridge to inspect Cadence/PSTX projects, collect evidence, query nets/refdes/tables/topology/datasheets, and avoid reading internal source code directly. Trae must call the HTTP bridge instead of assuming local Python.
---

# PSTX Bridge Analysis Skill for Trae

This skill is written for Trae or another external coding/analysis agent. Trae may run on a different computer from the PSTX project and Python environment, so Trae must use the HTTP Bridge protocol below instead of running `python`, importing `pstx_*`, scraping Web UI HTML, or guessing report table structures.

## Mental Model

- The **analysis/upper machine** owns the project files, Python environment, datasheet index, and PSTX parser.
- The **Trae machine** reads this skill and sends JSON requests to the upper machine Bridge.
- The Bridge exposes only whitelisted PSTX commands. It does not execute arbitrary shell commands.
- The Bridge command payload maps to the stable CLI schema, but Trae should treat it as HTTP JSON, not as local command-line execution.

## Bridge Setup

Trae must assume the bridge is already provided by the analysis/upper machine. In the current project, the bridge starts automatically when the operator starts the project Web UI or desktop UI on that machine. The default bridge port is the high, project-specific port `48765`.

Trae must not start the bridge itself and must not run Python. If the bridge is unavailable, report that the upper-machine PSTX Bridge is not connected and ask the operator to start the project service.

For Trae on another computer, the operator may expose the already-running bridge through VPN, SSH tunnel, reverse proxy, or a controlled LAN address. Trae should only consume the final URL/token.

Trae should be given:

- `PSTX_BRIDGE_URL`, for example `http://127.0.0.1:48765` or `http://analysis-host:48765`.
- Optional `PSTX_BRIDGE_TOKEN`; if provided, send it in `X-PSTX-Bridge-Token`.

If the bridge is unavailable, Trae should report that the upper-machine analysis service is not connected. Do not fall back to local Python unless the user explicitly says Trae is running on the analysis machine and local CLI execution is allowed.

Operator-only note, not a Trae action: the bridge is normally auto-started by `pstx_web.py` / `pstx_local_ui.py`; standalone bridge startup exists only for deployment debugging.

## HTTP Protocol

### Health

```bash
curl -s "$PSTX_BRIDGE_URL/v1/health" \
  -H "X-PSTX-Bridge-Token: $PSTX_BRIDGE_TOKEN"
```

Expected top-level fields:

- `ok`
- `schema_version=pstx-trae-bridge.v1`
- `status`
- `capability_count`

### Discover Capabilities

```bash
curl -s "$PSTX_BRIDGE_URL/v1/capabilities" \
  -H "X-PSTX-Bridge-Token: $PSTX_BRIDGE_TOKEN"
```

### Read Schema

```bash
curl -s "$PSTX_BRIDGE_URL/v1/schema" \
  -H "X-PSTX-Bridge-Token: $PSTX_BRIDGE_TOKEN"
```

Read one command schema:

```bash
curl -s "$PSTX_BRIDGE_URL/v1/schema/net-catalog" \
  -H "X-PSTX-Bridge-Token: $PSTX_BRIDGE_TOKEN"
```

### Discover Web-Analyzed Projects

Before asking the operator for a local project path, Trae must check whether the project has already been analyzed in the Web UI:

```bash
curl -s "$PSTX_BRIDGE_URL/v1/projects" \
  -H "X-PSTX-Bridge-Token: $PSTX_BRIDGE_TOKEN"
```

Use the returned `projects[].run_id` with `POST /v1/run`. `run_id="latest"` is allowed only when the user has confirmed the latest Web run is the intended project. The bridge will translate that Web run into a temporary bundle cache on the analysis machine; Trae must not use the temporary CLI path from `bridge.cli_argv` as a durable reference.

Example using the current Web project without a local path:

```json
{
  "command": "evidence-pack",
  "args": {
    "run_id": "latest",
    "refdes": ["U46"],
    "table_id": ["chip_pin_rows"]
  }
}
```

Commands that can reuse Web runs are the commands that read `--bundle-cache-in`: `query`, `batch-query`, `module-review`, `report-table`, `report-aggregate`, `evidence-pack`, `net-catalog`, `topology-netlist`, `cadence-page`, `cadence-index`, `csa-geometry`, and `schematic-pdf-annotate`.

### Run A Whitelisted Command

Use `POST /v1/run` with a JSON body:

```json
{
  "command": "inspect",
  "args": {
    "project_root": "/path/to/project"
  }
}
```

`command` must be one of `/v1/capabilities` or `/v1/schema` advertised commands. `args` uses snake_case keys; the bridge converts them to the underlying public command flags on the analysis machine.

For project evidence commands, prefer `args.run_id` from `/v1/projects` over `project_root`. Ask for `project_root` only when `/v1/projects` is empty or the user explicitly wants to analyze a new project that is not loaded in Web.

When the operator does provide a path, Trae does not need to force them to locate `worklib/<main module>`. The upper-machine CLI accepts a direct project root, `packaged`, a project container such as `A/B`, `A/B/worklib`, `A/B/worklib/<main module>`, or a supported `.zip/.tar*` project archive. For a container, the CLI reads the single `.cpm` filename stem as the main module name. If an archive is present near the container, the CLI copies it to a local `output/project_snapshots/` snapshot and analyzes that extracted copy so SMB changes cannot affect the current run. Read `project.snapshot` or `summary.project_input_snapshot` to report what source was actually used.

The response is the normal PSTX JSON envelope plus:

```json
{
  "bridge": {
    "interface": "pstx-trae-bridge",
    "schema_version": "pstx-trae-bridge.v1",
    "cli_exit_code": 0,
    "cli_argv": ["inspect", "/path/to/project"],
    "transport": "http-json",
    "bundle_source": "web_run",
    "run_id": "run_xxxxxxxxxxxxxxxx",
    "project_name": "demo_project"
  }
}
```

Always check `ok` first. If `ok=false`, surface `error_code`, `error_message`, and the command that failed.

### Read Background Agent Runs

If the PSTX Web UI starts a long-running Harness Agent task, the operator may give Trae an `agent_run_id`. Trae must not inspect `agent_workspace/` directly; use the Bridge:

```json
{
  "command": "agent-run-status",
  "args": {
    "agent_run_id": "report_xxxxxxxxxxxxxxxx"
  }
}
```

For generated summaries or downloadable drafts:

```json
{
  "command": "agent-run-artifacts",
  "args": {
    "agent_run_id": "report_xxxxxxxxxxxxxxxx"
  }
}
```

For partial/final execution trace:

```json
{
  "command": "agent-run-trace",
  "args": {
    "agent_run_id": "report_xxxxxxxxxxxxxxxx"
  }
}
```

Status values are `queued`, `running`, `waiting_for_user`, `completed`, `failed`, `cancelled`, and `incomplete`. If status is `running`, `waiting_for_user`, or `incomplete`, do not invent the final result; report `current_phase`, `progress`, `next_actions`, and whether `can_continue` is true.

When `agent-run-status` returns `partial_trace`, treat it as the authoritative checkpoint summary. Use `agent-run-trace` when you need steps/tool calls/evidence ids; use `agent-run-artifacts` when you need `result.json`, `trace.json`, `evidence_cards.json`, `task_ledger.md`, `answer.md`, or `review_draft.md`. Text artifacts include bounded `content_preview` fields for remote reading; Trae must still not read `agent_workspace/` paths directly.

## Recommended Workflow

### 1. Discover Capabilities

```json
{
  "command": "schema",
  "args": {}
}
```

Use this before hard-coding arguments. The schema is the current contract.

### 2. Inspect The Project

```json
{
  "command": "inspect",
  "args": {
    "project_root": "/path/to/project"
  }
}
```

Read:

- `project`: normalized project root.
- `project.snapshot`: project input normalization metadata; if `enabled=true`, this run used a copied local archive snapshot.
- `files`: key PSTX file status.
- `page_sources`: module order and `sch_1/page*.csv|csa` counts.
- `suggested_workflow`: safe follow-up commands.

If required files are missing, stop and explain what is missing instead of parsing source files manually.

### 3. Analyze Once And Cache

```json
{
  "command": "analyze",
  "args": {
    "project_root": "/path/to/project",
    "bundle_cache_out": "out/bundle-cache.json",
    "report_json_out": "out/report.json"
  }
}
```

Reuse `out/bundle-cache.json` for follow-up calls. This avoids repeated heavy parsing on the upper machine.

### 4. Collect Mixed Evidence

Use `evidence-pack` when a question mentions mixed targets such as refdes + net + HQ + table.

```json
{
  "command": "evidence-pack",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "refdes": ["U46", "PU12"],
    "net": ["P3V3"],
    "hq": ["HQ11112042009"],
    "page": ["131", "152"],
    "table_id": ["chip_pin_rows"]
  }
}
```

Read:

- `evidence_pack.target_summary`: requested and found target counts.
- `evidence_pack.items`: per-target evidence.
- `evidence_pack.tables`: report-table previews with truncation metadata.
- `evidence_pack.recommended_next_commands`: safe follow-up route.

If a table preview is truncated, page it through the bridge:

```json
{
  "command": "report-table",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "table_id": "chip_pin_rows",
    "offset": 0,
    "limit": 200
  }
}
```

### 5. Query Homogeneous Targets In Bulk

```json
{
  "command": "batch-query",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "mode": "位号",
    "items": "U1,U2,U46"
  }
}
```

Common modes include `位号`, `网络`, `HQ料号`, and `页码`.

### 6. Discover Net Labels Before Deep Evidence

Use `net-catalog` when the user mentions fuzzy nets, interface abbreviations, power rails, differential pairs, or asks “有哪些网标/网络”.

```json
{
  "command": "net-catalog",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "query": "PCE",
    "kind": "differential",
    "include_nodes": true,
    "limit": 200
  }
}
```

Useful `kind` values:

- `power`: power rails such as `P3V3`, `VDD`, `VBAT`.
- `ground`: `GND`, `AGND`, `PGND`, `VSS`.
- `differential`: PCIe/USB/MIPI/LVDS style P/N nets.
- `unnamed`: generated or unnamed nets that need page/location follow-up.
- `signal`: ordinary named signal nets.

Rules:

- Do not guess exact net labels from memory. If unsure, run `net-catalog` first.
- Treat `PCE`, `P5E`, `PCI-E`, and similar aliases through `business-dictionary`, not ad-hoc guessing.
- Do not treat `net-catalog` samples as complete proof; use `detail_command`, `evidence-pack`, `report-table`, or `topology-netlist` for conclusions.
- If `truncated=true`, page with `offset/limit` before saying no more nets exist.

### 7. Use Aggregation For Counts

Never count unique values from a truncated table preview. Use aggregation.

```json
{
  "command": "report-aggregate",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "table_id": "page_rows",
    "column": "页码",
    "operation": "top"
  }
}
```

Use this for questions like:

- “有多少页原理图？”
- “哪些页出现某类问题？”
- “某列有哪些唯一值？”

### 8. Review By Module

```json
{
  "command": "module-review",
  "args": {
    "project_root": "/path/to/project",
    "module_type": "子模块",
    "module_name": "i2c_repeater_9617_cbb_v3"
  }
}
```

Use module review when the user asks to separate main module and submodule review.

### 9. Compare Projects

`left_project_root` and `right_project_root` use the same project input normalization as `analyze`: each side may be a direct project root, project container, worklib directory, or supported archive. If a side resolves through a copied snapshot, the compare result's side metadata will point at the normalized local project root.

```json
{
  "command": "compare",
  "args": {
    "left_project_root": "/path/to/left/project",
    "right_project_root": "/path/to/right/project",
    "detail_limit": 1000
  }
}
```

### 10. Use Business Dictionary And Topology

Read business terms before interpreting project abbreviations or interface names:

```json
{
  "command": "business-dictionary",
  "args": {}
}
```

Then export a lightweight topology summary:

```json
{
  "command": "topology-netlist",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "stdout": "summary",
    "view": "summary",
    "supply_mode": "grouped",
    "supply_limit": 12
  }
}
```

Summary topology is intentionally visual/LLM-light: PMIC supply fanout is returned as `supply_edge_groups[]` plus a few `supply_edges[]` samples while `counts.total_supply_edge_count` keeps the full total. Do not treat the samples as complete supply evidence.

Use focused full topology only when the user asks for chip-level connection details:

```json
{
  "command": "topology-netlist",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "focus_refdes": "U46",
    "stdout": "full",
    "view": "full",
    "supply_mode": "details",
    "include_connectors": true
  }
}
```

Rules:

- `topology_business_view.review_queue` tells what to review first; it is not the complete graph.
- `supply_edge_groups` is a grouped supply overview for speed/readability; use full/details, `query_llm_topology_netlist`, or `get_llm_topology_edge` before making a claim about every load on a rail.
- `topology_cache_status` tells whether the derived topology came from local `output/analysis_cache`.
- `topology_netlist.review_tasks` is the more Agent-friendly queue. Use each task's `detail_tool` semantics to decide whether to fetch node/edge details before making a review claim.
- `nodes[].llm_device_identity_hint` packages PART_NAME / CDS_PART_NAME / spec / HQ evidence for server-hardware device-role judgment. Treat it as a hint, not a deterministic regex classification.
- `edges[].interface_completeness` shows observed and missing key sub-signals such as PCIe TX/RX/REFCLK/PERST or I2C SCL/SDA. Missing items mean “needs detail review”, not automatic failure.
- If topology output says `truncated=true`, run a focused topology request or write/read the full artifact with `out`.
- Cite `llm_topology_node`, `llm_topology_edge`, `llm_topology_supply_edge`, or `llm_topology_review_task` evidence ids when making topology conclusions.

### 11. Inspect Raw Cadence Page Semantics

Use `cadence-page` when the user asks about a specific schematic page's raw Cadence connectivity evidence, such as WIRE/DOT/SIG_NAME, network labels, ports, off-page connectors, bus names, No Connect markers, or unknown CSA rows.

```json
{
  "command": "cadence-page",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "page": 114,
    "stdout": "objects",
    "limit": 200
  }
}
```

For a single object detail:

```json
{
  "command": "cadence-page",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "page": 114,
    "object_id": "p114-net_label-1"
  }
}
```

Rules:

- `cadence-page` is raw page evidence. It does not replace `topology-netlist` for chip-level connection conclusions.
- `connectivity_summary.unbound_semantics` means the semantic object was visible on the page but was not geometrically bound to a wire. Do not cite it as proof of connection.
- Use `stdout=summary` for quick page counts, `stdout=objects` for object-level evidence, and `stdout=full` when you need connectivity groups and object lists together.
- If `truncated=true`, repeat with a higher `limit` or focus with `object_id` before concluding an object is absent.

### 12. Inspect Project-Level Cadence Semantic Index

Use `cadence-index` when the user asks where a Cadence net label, port, off-page connector, bus name, No Connect marker, or unbound page semantic appears across the project.

```json
{
  "command": "cadence-index",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "stdout": "full",
    "query": "P1V8",
    "kind": "all",
    "limit": 200
  }
}
```

Rules:

- `cadence-index` is a Cadence page-graphic semantic catalog. `net-catalog` is the PSTX netlist catalog; exact name matches are hints only.
- `offpage_link_rows` are same-name page evidence, not a complete electrical connection claim.
- `unbound_semantic_rows` must be cited as visible-but-unbound evidence, not as connected evidence.

### 13. Scan CSA Geometry Checks

Use `csa-geometry` when the user asks for DE HDL CSA geometric review evidence: DOT four-way crosses, CIRCLE marks, optional ARC-fitted circle candidates, missing page numbers, or package-style CSV/JSON/HTML outputs.

```json
{
  "command": "csa-geometry",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "stdout": "hits",
    "limit": 200
  }
}
```

For geometric findings with same-page connectivity evidence:

```json
{
  "command": "csa-geometry",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "include_connectivity": true,
    "page": 114,
    "stdout": "full",
    "limit": 200
  }
}
```

For recursive worklib scan and package-style report files:

```json
{
  "command": "csa-geometry",
  "args": {
    "project_root": "/path/to/worklib",
    "recursive": true,
    "workers": 8,
    "executor": "thread",
    "check_missing": true,
    "out_dir": "out/csa-geometry",
    "json": true,
    "html": true
  }
}
```

Rules:

- `csa-geometry` is geometry evidence. `include_connectivity=true` adds a conservative `semantic_overlay` with page-level signal labels, ports, off-page connectors, bus names, and No Connect evidence; it still does not prove an electrical short or replace `cadence-page` / `topology-netlist`.
- DOT/circle rows include source-line evidence such as original CSA lines and nearby context; use those fields when quoting why a finding exists.
- T junctions, dotless visual crosses, and diagonal lines are intentionally not reported as DOT four-way crosses.
- CLI defaults to CIRCLE only; pass `include_arcs=true` when ARC-fitted circle candidates are needed, and treat ARC results as manual-review candidates.
- `fail_on_findings` and `fail_on_circles` can be used by automation; a bridge caller should still inspect the returned JSON payload.

### 14. Build Schematic PDF Annotation Overlays

Use `schematic-pdf-annotate` when the user wants review findings highlighted on a schematic PDF: compare added/removed parts, BOM warnings, derating warnings, CSA findings, or manual page markers.

```json
{
  "command": "schematic-pdf-annotate",
  "args": {
    "run_id": "latest",
    "pdf": "/path/on/analysis-machine/schematic.pdf",
    "refdes": ["U46", "R120"],
    "pdf_page_map_json": "{\"PAGE114\": 1}",
    "stdout": "full",
    "limit": 200
  }
}
```

For a direct coordinate target that is already in PDF coordinate space:

```json
{
  "command": "schematic-pdf-annotate",
  "args": {
    "run_id": "latest",
    "pdf": "/path/on/analysis-machine/schematic.pdf",
    "target_json": "{\"kind\":\"coordinate\",\"page\":\"PAGE114\",\"pdf_page_number\":1,\"pdf_bbox\":[10,20,80,60],\"label\":\"降额提醒\",\"severity\":\"warning\"}",
    "stdout": "full"
  }
}
```

Rules:

- The PDF path must be on the analysis machine unless the Web API multipart upload is used by the UI. Trae must not assume it can upload local files through the Bridge.
- `project_page` in the returned annotations is the user-visible real schematic page after submodule/module_order mapping. Do not replace it with a submodule-local page.
- Always prefer an explicit `pdf_page_map_json` when the PDF page order may differ from project `PAGE<N>` labels. The CLI can also use a unique PDF text `PAGE<N>` label as evidence.
- When storing/reusing a PDF map, prefer `{"pdf_sha256":"...","pages":{"PAGE114":1}}`; if the hash no longer matches the current PDF, PSTX rejects the map and reports a warning.
- Do not enable `allow_page_number_fallback` unless the user explicitly confirms the PDF is exported in exactly the same order as project real pages; fallback hits are marked `page_label_number_weak`.
- A target with `pdf_bbox` is drawable only when it also has a reliable PDF page from `pdf_page_number`, `pdf_page_map_json`, or a unique PDF text page label.
- Do not treat raw PSTX/Cadence `XY` as PDF coordinates. It can only become a drawable bbox when `page_calibrations` are provided; otherwise the payload should be reported as page-level evidence.
- `confidence` is authoritative: `explicit_pdf_bbox`, `calibrated_xy`, and `pdf_text_match` are drawable; `page_only` should be shown as a page-side note; `unmatched` needs manual mapping.
- `pdf_text_match` depends on PDF text bboxes. If the PDF is plotted as outlines/images, the result may fall back to page-only evidence.

### 15. Use Datasheet Templates And Parameter Cards

Before asking an LLM to read datasheet evidence, load the review template. It tells the model what fields to extract and what schematic evidence must be checked.

Shared skill source: if the PSTX repository files are available to Trae, first read `harness_skills/datasheet-key-info/SKILL.md`. If Trae only has Bridge access, fetch the same card through `harness-skills datasheet-key-info --include-body`. It is the same MinerU / datasheet key-information playbook that the Web Harness Agent can load, so Trae and the in-project Agent should follow the same evidence order and output shape.

```json
{
  "command": "harness-skills",
  "args": {
    "skill_id": "datasheet-key-info",
    "include_body": true
  }
}
```

The upper machine indexes datasheet PDFs with MinerU by default. Treat `datasheet-status.extractor.mode=mineru` as the normal path. If MinerU is unavailable or a document is marked `needs_manual_review`, report that PDF extraction did not produce reliable evidence; do not silently treat pypdf/fallback snippets as equivalent unless the operator explicitly configured `PSTX_PDF_EXTRACTOR=auto` or `pypdf`.

```json
{
  "command": "datasheet-template",
  "args": {
    "template_id": "complex_chip"
  }
}
```

Check datasheet index status:

```json
{
  "command": "datasheet-status",
  "args": {
    "include_documents": true,
    "limit": 200
  }
}
```

Search datasheet chunks for candidate evidence:

```json
{
  "command": "datasheet-search",
  "args": {
    "query": "recommended operating VDD power sequence",
    "limit": 10
  }
}
```

Search deterministic parameter cards for numeric facts:

```json
{
  "command": "datasheet-parameters",
  "args": {
    "query": "VDD voltage current thermal sequence",
    "limit": 20
  }
}
```

Complex-chip playbook, based on the prior 64144 datasheet review:

1. Start with `datasheet-template complex_chip`; use its sections as the answer outline.
2. Run `datasheet-status` with `include_documents=true`; confirm the target document is indexed by MinerU, has nonzero `page_count`, `chunk_count`, and `parameter_count`, and is not `needs_manual_review`.
3. Identify the right document by HQ code, part name, ordering/package text, and revision; if multiple documents match, keep them separate until identity is proven.
4. For large chips like the 64144 case, search these evidence groups before summarizing: `recommended operating conditions`, `absolute maximum ratings`, `power rail voltage`, `power consumption current`, `power up sequence`, `power down sequence`, `reset timing`, `pin description voltage domain`, `IO threshold`, `clock requirements`, `strap boot mode`, `thermal characteristics`, and `junction temperature`.
5. Use `datasheet-parameters` first for numeric facts. Good first filters/queries are `power_rail_voltage`, `power_budget_current`, `power_sequence_timing`, `thermal_characteristic`, `junction_temperature_limit`, `recommended operating`, `absolute maximum`, and the concrete rail/interface names from the schematic.
6. For every high-risk numeric conclusion, read the detail evidence through the harness/Web detail tool when available (`get_datasheet_parameter`, `get_datasheet_chunk`, or `get_datasheet_page_excerpt`). Bridge-only callers cannot invoke those Harness detail tools directly; in Bridge-only mode, cite the full `datasheet-parameters` card fields (`parameter_id`, `doc_id`, page, chunk locator, value, condition, and source_text`) and mark any missing detail as an evidence gap for the Web Harness/operator.
7. Map each datasheet fact back to schematic evidence with `evidence-pack`, `batch-query`, `topology-netlist`, `cadence-index`, or `cadence-page`. For example, a rail voltage requirement must cite the rail net/topology evidence; a sequence requirement must cite EN/PGOOD/RESET connections; an interface voltage-domain claim must cite both sides of the interface.
8. Output three buckets: confirmed datasheet facts, schematic evidence that matches them, and missing/manual-review items. Do not collapse unknowns into pass/fail.

Rules:

- `datasheet-search` is a locator, not proof. Do not answer voltage/current/temperature/timing questions only from a snippet.
- The shared Harness skill `datasheet-key-info` is canonical for output shape: `confirmed datasheet facts`, `schematic evidence mapped to those facts`, and `missing/manual-review items`.
- Prefer `datasheet-parameters` for numeric facts and then cite `parameter_id`, `doc_id`, page, and chunk locator. For 64144-style complex chips, parameter cards are the fastest way to avoid losing table values across PDF chunk boundaries.
- Map datasheet facts back to schematic evidence using `evidence-pack`, `batch-query`, and `topology-netlist`.
- Use `cadence-index` / `cadence-page` when the question depends on raw page labels, ports, off-page links, Bus names, or No Connect evidence.
- If `datasheet-status.configured=false`, say the upper-machine datasheet source is unavailable and ask the operator to configure/reindex it.
- If `datasheet-status.failures` or a document `error` mentions MinerU, tell the operator which PDF failed and whether they should install/configure MinerU, reindex, or intentionally switch to `PSTX_PDF_EXTRACTOR=auto` for a text-only fallback.

## Answering Rules For Trae

- Cite the bridge command and evidence fields used for important conclusions.
- If evidence is incomplete, say what bridge request should be run next.
- If the user asks a broad question, start with `inspect`, `analyze`, then `evidence-pack`.
- If the user asks for exact counts, use `report-aggregate`, not preview rows.
- If the user asks for detailed rows, use `report-table` with pagination.
- If the user asks about many refdes or exact nets, use batch tools.
- If the user asks about fuzzy nets, interface aliases, power rails, differential nets, or unnamed nets, use `net-catalog` before detail evidence.
- If the bridge returns `ok=false`, do not fabricate an answer; surface the structured error and propose the next request.

## Typical Chinese Prompts And Bridge Requests

用户：“帮我看 U46 的料号、页码和 pin/net。”

```json
{
  "command": "evidence-pack",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "refdes": ["U46"],
    "table_id": ["chip_pin_rows"]
  }
}
```

用户：“这个项目有多少页原理图？”

```json
{
  "command": "report-aggregate",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "table_id": "page_rows",
    "column": "页码"
  }
}
```

用户：“项目里 PCE/PCIe 相关的差分网有哪些？”

```json
{
  "command": "net-catalog",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "query": "PCE",
    "kind": "differential",
    "include_nodes": true
  }
}
```

用户：“有哪些未命名网络需要定位？”

```json
{
  "command": "net-catalog",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "kind": "unnamed",
    "include_nodes": true
  }
}
```

用户：“看一下第 114 页的 Cadence 原始连接标注。”

```json
{
  "command": "cadence-page",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "page": 114,
    "stdout": "objects"
  }
}
```

用户：“扫描一下 CSA 几何规范候选，导出十字和画圈明细。”

```json
{
  "command": "csa-geometry",
  "args": {
    "bundle_cache_in": "out/bundle-cache.json",
    "stdout": "hits",
    "out_dir": "out/csa-geometry",
    "json": true
  }
}
```

用户：“对比两个项目里芯片和连接器的差异。”

```json
{
  "command": "compare",
  "args": {
    "left_project_root": "/path/to/A",
    "right_project_root": "/path/to/B",
    "detail_limit": 1000
  }
}
```

用户：“电脑 A 上准备一个给无外网电脑 B 的完整迁移包，包含 Python 便携包、MinerU 和 wheelhouse。”

```json
{
  "command": "offline-migration",
  "args": {
    "offline_action": "prepare",
    "out_dir": "output/offline_migration",
    "name": "dehdl-b",
    "asset_cache_dir": "output/offline_migration/_asset_cache",
    "target_platform": "windows-amd64",
    "target_profile": "windows-rtx4060-cuda",
    "python_version": "3.10.11",
    "python_mirror": "tuna",
    "download_mineru_models": true,
    "mineru_model_source": "huggingface",
    "mineru_model_type": "pipeline",
    "huggingface_endpoint": "https://hf-mirror.com",
    "download_wheels": true,
    "pip_index_url": "https://pypi.tuna.tsinghua.edu.cn/simple",
    "pip_extra_index_url": "https://download.pytorch.org/whl/cu121",
    "include_mineru_wheels": true,
    "mineru_wheel_spec": "mineru[pipeline]",
    "no_zip": true
  }
}
```

用户：“电脑 B 上检查迁移包有没有丢文件。”

```json
{
  "command": "offline-migration",
  "args": {
    "offline_action": "verify",
    "package_root": "D:/dehdl-b"
  }
}
```

Offline migration rules:

- `offline-migration prepare` is an operator-triggered packaging action for the upper machine. Do not run it unless the user explicitly asks to prepare a migration package.
- `prepare` may download Python archives or wheels from mirrors on computer A. `verify` is offline-only and must not download anything on computer B.
- Prefer folder-first migration: the operator scripts `scripts/PREPARE_MIGRATION_A.cmd` / `.ps1` now default to `--no-zip`, so the user can manually compress the generated bundle folder with a tool that handles large Python/MinerU/model assets better. Use script `-MakeZip` or Bridge/CLI `no_zip=false` only if the user explicitly wants the script-created zip.
- Keep `asset_cache_dir` enabled unless the operator explicitly requests a clean rebuild. The default `<out_dir>/_asset_cache` reuses portable Python archives, MinerU pipeline models/config, and wheelhouse assets; when requirements change, it seeds the old wheels first and then downloads only missing or changed dependencies. Use `no_reuse_assets=true` only to diagnose cache pollution or force a full refresh.
- Computer B must not be assumed to have system Python. `prepare` requires a portable Python source by default; only use `allow_system_python_on_b=true` when the operator explicitly says B already has Python.
- Prefer `python_mirror=tuna` or an operator-provided `python_mirror_base`/`python_url` for China/offline staging. The default Python filename targets Windows embeddable Python: `python-<version>-embed-amd64.zip`.
- For RTX 4060 Windows B, use `target_profile=windows-rtx4060-cuda`. Prefer a tested CUDA-capable MinerU venv with `mineru_venv` when the operator has one; if it is omitted, `prepare` will auto-detect project `.venv-mineru`, or create it on computer A and install the default `mineru[pipeline]` when MinerU models/config are being prepared.
- If the operator does not already have a local model directory, use `download_mineru_models=true`, `mineru_model_source=huggingface`, `mineru_model_type=pipeline`, and `huggingface_endpoint=https://hf-mirror.com` on computer A. This invokes MinerU's own `mineru-models-download` and then packages the generated local model directory/config.
- `include_mineru_wheels=true` downloads MinerU only as a backup reinstall path. Use the default `mineru_wheel_spec=mineru[pipeline]` because PSTX calls `mineru -b pipeline`; do not use `mineru[all]` unless the operator explicitly needs extra VLM/LMDeploy backends. If `mineru_venv` is included, MinerU wheel resolution failure should be reported as a warning rather than blocking the migration; use `strict_mineru_wheels=true` only when the user explicitly requires a complete MinerU wheelhouse.
- The repository also provides operator scripts: `scripts/PREPARE_MIGRATION_A.cmd` / `.ps1` for computer A; generated bundles provide `RUN_SETUP_B.bat/.ps1`, `RUN_VERIFY_B.*`, `RUN_INSTALL_WHEELHOUSE_B.*`, and `START_WEB_B.*` for computer B.
- The generated package contains standard-library-only `RUN_SETUP_B.*`, `RUN_VERIFY_B.*` / `VERIFY_OFFLINE_B.py`; prefer those on computer B before using the project CLI, because project dependencies may not be installed yet.
- If verification reports missing runtime imports and `wheelhouse/` exists, tell the operator to run `RUN_INSTALL_WHEELHOUSE_B.*` once and verify again.
- If `verification.ok=false`, report the `issues[]` paths exactly and do not claim the target machine is ready.

## Do Not Do

- Do not assume Trae and Python are on the same machine.
- Do not run `python`, `python3`, `pstx_cli.py`, or shell commands from Trae unless the user explicitly says local CLI execution is allowed.
- Do not import `pstx_*` modules.
- Do not scrape `http://localhost` report pages for data.
- Do not read `trash/**`, `unused_code/**`, or archived docs unless the user asks for history.
- Do not assume a preview contains all rows.
- Do not call Web-only APIs when the same bridge command exists.
- Do not modify PSTX/Cadence project files through this skill.
