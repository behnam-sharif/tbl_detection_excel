"""
pipeline.py — Top-level orchestrator for the table detection pipeline.

Provides:
  - process_single_sheet(): full pipeline for one sheet
  - parse_workbook(): parallel processing across sheets via ThreadPoolExecutor
"""

from __future__ import annotations
import logging
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Tuple, Dict, Optional

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from parse_tools.models import BBox, BlockCandidate, TableDecision
from parse_tools.sheet_accessor import SheetAccessor
from parse_tools.presence import get_presence_matrix_from_accessor
from block_generation import generate_block_candidates
from table_framing_heuristics import run_framing_pipeline

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# QC checks
# ---------------------------------------------------------------------------

def _qc_check_block(
    block: BlockCandidate, decision: TableDecision
) -> List[str]:
    """
    Validate a (block, decision) pair. Returns list of warning strings (empty = pass).
    """
    warnings: List[str] = []
    bbox = block.bbox

    # 1. bbox coords non-negative, r1 <= r2, c1 <= c2
    if bbox.r1 < 0 or bbox.c1 < 0:
        warnings.append(f"negative bbox coords: {bbox}")
    if bbox.r1 > bbox.r2 or bbox.c1 > bbox.c2:
        warnings.append(f"inverted bbox: {bbox}")

    # 2. shrink -> final_bbox must differ from original_bbox
    if decision.action == "shrink":
        if decision.final_bbox == decision.original_bbox:
            warnings.append("shrink decision but final_bbox == original_bbox")

    # 3. keep/shrink -> final_bbox must not be None
    if decision.action in ("keep", "shrink") and decision.final_bbox is None:
        warnings.append(f"{decision.action} decision but final_bbox is None")

    # 4. reject -> final_bbox should be None
    if decision.action == "reject" and decision.final_bbox is not None:
        warnings.append("reject decision but final_bbox is not None")

    # 5. confidence in [0.0, 1.0]
    if not (0.0 <= decision.confidence <= 1.0):
        warnings.append(f"confidence {decision.confidence} out of [0, 1]")

    # 6. n_cells_on > 0 for non-reject decisions
    if decision.action != "reject" and block.n_cells_on <= 0:
        warnings.append(f"non-reject with n_cells_on={block.n_cells_on}")

    return warnings


# ---------------------------------------------------------------------------
# Single-sheet pipeline
# ---------------------------------------------------------------------------

def process_single_sheet(
    file_path: str,
    sheet_name: str,
    bridging_rows: int = 2,
    bridging_cols: int = 1,
    connectivity: int = 4,
    min_on_cells: int = 4,
    do_bridging: bool = True,
    max_framing_passes: int = 2,
) -> List[Tuple[BlockCandidate, TableDecision]]:
    """
    Full pipeline for one sheet:
      SheetAccessor -> presence matrix -> generate_block_candidates
      -> run_framing_pipeline -> sorted results

    Catches exceptions -> logs warning -> returns [].
    """
    t0 = time.perf_counter()
    try:
        with SheetAccessor(file_path, sheet_name) as accessor:
            logger.info(
                "Loaded sheet '%s' (%d rows x %d cols, %d merged coords)",
                sheet_name, accessor.nrows, accessor.ncols,
                len(accessor.merged_cell_coords),
            )

            # Build presence matrix
            presence = get_presence_matrix_from_accessor(accessor)

            # Generate block candidates
            t1 = time.perf_counter()
            candidates = generate_block_candidates(
                sheet_id=sheet_name,
                presence_matrix=presence,
                do_bridging=do_bridging,
                bridging_rows=bridging_rows,
                bridging_cols=bridging_cols,
                connectivity=connectivity,
                min_on_cells=min_on_cells,
            )
            t2 = time.perf_counter()
            logger.info(
                "Sheet '%s': %d blocks generated in %dms",
                sheet_name, len(candidates), int((t2 - t1) * 1000),
            )

            # Run framing heuristics
            results = run_framing_pipeline(
                blocks=candidates,
                accessor=accessor,
                presence_matrix=presence,
                base_row=1,
                base_col=1,
                max_passes=max_framing_passes,
            )

            # Log summary
            action_counts: Dict[str, int] = {}
            for _, decision in results:
                action_counts[decision.action] = action_counts.get(decision.action, 0) + 1
            summary = ", ".join(f"{v} {k}" for k, v in sorted(action_counts.items()))
            tables_detected = sum(1 for _, d in results if d.action in ("keep", "shrink"))

            logger.info(
                "Sheet '%s': framing -> %s",
                sheet_name, summary,
            )

            # QC checks
            for block, decision in results:
                qc_warnings = _qc_check_block(block, decision)
                for w in qc_warnings:
                    logger.warning("QC [%s] block %s: %s", sheet_name, block.bbox, w)

                logger.debug(
                    "Block %s (density=%.2f, border_coverage=%.2f) -> %s (confidence=%.2f)",
                    block.bbox, block.density,
                    0.0,  # border_coverage not stored on block; logged in framing
                    decision.action, decision.confidence,
                )

            elapsed = int((time.perf_counter() - t0) * 1000)
            logger.info(
                "Sheet '%s': complete in %dms (%d tables detected)",
                sheet_name, elapsed, tables_detected,
            )

            return results

    except Exception as e:
        logger.warning("Error processing sheet '%s': %s", sheet_name, e)
        return []


# ---------------------------------------------------------------------------
# Workbook-level pipeline
# ---------------------------------------------------------------------------

def parse_workbook(
    file_path: str,
    sheet_names: Optional[List[str]] = None,
    max_workers: int = 4,
    **pipeline_kwargs,
) -> Dict[str, List[Tuple[BlockCandidate, TableDecision]]]:
    """
    Top-level entry point. Processes all (or selected) sheets in parallel.

    Each thread creates its own SheetAccessor (no shared state).
    Results ordered by original sheet order.
    Excludes "Shape_Info" sheets by default.
    """
    # Discover sheets if not provided
    if sheet_names is None:
        wb = load_workbook(file_path, data_only=True)
        sheet_names = [n for n in wb.sheetnames if n != "Shape_Info"]
        wb.close()

    if not sheet_names:
        return {}

    effective_workers = min(max_workers, len(sheet_names))
    results: Dict[str, List[Tuple[BlockCandidate, TableDecision]]] = {}

    logger.info(
        "parse_workbook: %d sheets, %d workers, file=%s",
        len(sheet_names), effective_workers, file_path,
    )

    with ThreadPoolExecutor(max_workers=effective_workers) as executor:
        future_to_sheet = {
            executor.submit(
                process_single_sheet, file_path, sn, **pipeline_kwargs
            ): sn
            for sn in sheet_names
        }

        for future in as_completed(future_to_sheet):
            sn = future_to_sheet[future]
            try:
                results[sn] = future.result()
            except Exception as e:
                logger.warning("Thread error for sheet '%s': %s", sn, e)
                results[sn] = []

    # Return in original sheet order
    ordered: Dict[str, List[Tuple[BlockCandidate, TableDecision]]] = {}
    for sn in sheet_names:
        ordered[sn] = results.get(sn, [])

    return ordered
