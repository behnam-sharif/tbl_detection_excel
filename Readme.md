# xl_parse

Excel table detection and parsing: identify and extract table boundaries from spreadsheets using a multi-stage heuristic pipeline. Handles complex layouts with multiple tables per sheet. Output is suitable for downstream ML/LLM use.

## Pipeline overview

The pipeline runs **per sheet** and produces final table candidates (bounding boxes, header rows, structural type, scores) for later ML or LLM processing.

![Per-sheet pipeline: load grid → presence matrix → block generation → cell classification → table framing → grammar splitting → QC gate → final candidates](assets/per-sheet-pipeline.png)

## Project Overview

xl_parse is an Excel table detection and parsing system. It identifies and extracts table boundaries from spreadsheets using a multi-stage heuristic pipeline, handling complex layouts with multiple tables per sheet.

## Per-sheet pipeline summary (parallelized)

1. **Load sheet** — `SheetAccessor` opens the workbook/sheet and exposes values, formulas, and styles.
2. **Presence matrix** — Build a 0/1 matrix (`1` for non-empty cells, `0` for empty cells).
3. **Block generation** — Find connected components, optionally bridge small gaps, then expand context to include likely header/label regions.
4. **Framing heuristics** — For each block, classify cells and features, then decide `keep`/`reject`/`shrink` (and sometimes split large rejects). Shrunk blocks are re-evaluated.
5. **Grammar heuristics** — For kept/shrunk blocks, detect multi-table boundaries using separator lines, header restarts, and change-points.
6. **Post-split QC gate** — Apply pullback, fast reject filters, tiny-fragment merges, and refinement passes (re-split/adjacent merge).
7. **Final candidates** — Build `FinalCandidate` objects with bbox, header/key-column info, structural type, scores, and QC flags.
8. **Per-sheet output** — Return `SheetResult` with `results`, `final_candidates`, `qc_events`, and `cut_candidates`.
9. **Export helpers** — `ml_export.py` provides optional flattening for ML rows, cut candidates, and LLM-QC payload stubs.

## Dependencies

Install with: `pip install -r requirements.txt` (openpyxl, streamlit)

No test framework, linter, or build system is configured.

## Architecture

### Pipeline Flow

```
Excel File → SheetAccessor → Presence Matrix → Block Generation (connected components)
    → Table Framing (heuristic evaluation) → Grammar Splitting → Export
```

### Key Modules

- **`pipeline.py`** — Top-level orchestrator. `parse_workbook()` processes all sheets in parallel via ThreadPoolExecutor; `process_single_sheet()` runs the full pipeline for one sheet.
- **`parse_tools/models.py`** — All dataclasses live here to avoid circular imports: `BBox`, `BlockCandidate`, `CellFeatures`, `BlockFeatures`, `TableDecision`, `CutCandidate`, `Config`. The `Config` class centralizes all thresholds.
- **`parse_tools/__init__.py`** — Re-exports public APIs from submodules.
- **`parse_tools/sheet_accessor.py`** — Thread-safe openpyxl wrapper loading dual workbooks (one for formulas, one for cached values). Each thread creates its own instance.
- **`parse_tools/presence.py`** — Builds binary presence matrix (has content vs empty).
- **`parse_tools/cell_classifier.py`** — Cell-level type detection (formula → blank → number → text priority) and style extraction.
- **`parse_tools/feature_extraction.py`** — Aggregates per-cell features into block-level features (density, type histograms, gaps, bold%, fill%, border coverage).
- **`block_generation.py`** — Connected components analysis with optional gap-bridging and context expansion (header rows above, label columns left).
- **`table_framing_heuristics.py`** — Multi-pass heuristic evaluation. Classifies blocks as keep/reject/shrink/expand based on density, borders, headers. Applies shrinking and splitting.
- **`grammar_heuristics.py`** — Post-framing multi-table splitting using three strategies: separator detection, header restart detection, and text/numeric change-point detection. Candidates are scored independently per strategy and merged.
- **`ml_export.py`** — Exports flat dicts for ML training rows and cut candidates. Contains a stub for LLM-based QC payloads.

### Design Conventions

- **BBoxes are 0-indexed inclusive** internally; converted to Excel 1-indexed ranges (e.g., "A1:D10") for display via `bbox_to_excel_range()`.
- **Provenance tracking**: `BlockCandidate.provenance` records transformation history (e.g., "cc+bridge+expand+shrink").
- **Multi-pass refinement**: Framing runs up to 2 passes — evaluate, shrink/split, then re-evaluate changed blocks.
- **No shared mutable state**: Parallel sheet processing works because each thread owns its SheetAccessor.
- **All thresholds in `Config`**: No magic numbers scattered in code.
