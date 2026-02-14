"""
run_export.py — File I/O writer for run directory output.

Writes the complete set of output files expected by Repo 2:
  final_candidates.jsonl, features.csv, qc_events.jsonl,
  cut_scores.jsonl, and run_meta.json.
"""

from __future__ import annotations
import csv
import json
import logging
import os
from datetime import datetime, timezone
from typing import Dict, List, Optional

from parse_tools.models import BBox, Config, SheetResult
from parse_tools.id_utils import compute_workbook_id, compute_candidate_id
from final_candidates import export_final_candidate
from ml_export import export_cut_candidate, export_qc_event

logger = logging.getLogger(__name__)


def export_run(
    file_path: str,
    workbook_results: Dict[str, SheetResult],
    base_dir: str = "runs",
    config: Optional[Config] = None,
) -> str:
    """
    Write a complete run directory for Repo 2 consumption.

    Returns the path of the created output directory.
    """
    cfg = config or Config()

    # 1. Compute workbook ID
    workbook_id = compute_workbook_id(file_path)

    # 2. Create auto-named subdirectory
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    dir_name = f"{workbook_id[:8]}_{timestamp}"
    output_dir = os.path.join(base_dir, dir_name)
    os.makedirs(output_dir, exist_ok=True)

    # 3. Collect all data across sheets
    all_final_rows: List[dict] = []
    all_training_rows: List[dict] = []
    all_qc_rows: List[dict] = []
    all_cut_rows: List[dict] = []
    total_candidates = 0

    for sheet_name, sheet_result in workbook_results.items():
        # --- Final candidates: assign IDs and export ---
        for rank, fc in enumerate(sheet_result.final_candidates):
            cid = compute_candidate_id(workbook_id, sheet_name, fc.bbox, rank)
            fc.candidate_id = cid

            all_final_rows.append(
                export_final_candidate(fc, workbook_id=workbook_id, rank=rank)
            )
            total_candidates += 1

        # --- Features (pre-computed training rows from pipeline) ---
        # Build a bbox -> candidate_id map from final candidates
        cid_map: Dict[tuple, str] = {}
        for fc in sheet_result.final_candidates:
            key = (fc.sheet_name, fc.bbox.r1, fc.bbox.c1, fc.bbox.r2, fc.bbox.c2)
            cid_map[key] = fc.candidate_id

        for row in sheet_result.training_rows:
            row = dict(row)  # copy to avoid mutating original
            row["workbook_id"] = workbook_id
            bbox_key = (row["sheet_name"], row["bbox_r1"], row["bbox_c1"], row["bbox_r2"], row["bbox_c2"])
            row["candidate_id"] = cid_map.get(bbox_key, "")
            all_training_rows.append(row)

        # --- QC events ---
        for event in sheet_result.qc_events:
            row = export_qc_event(event)
            row["sheet_name"] = sheet_name
            row["workbook_id"] = workbook_id
            all_qc_rows.append(row)

        # --- Cut candidates ---
        for bbox_key, cuts in sheet_result.cut_candidates.items():
            parts = bbox_key.split(",")
            if len(parts) == 4:
                bbox = BBox(int(parts[0]), int(parts[1]), int(parts[2]), int(parts[3]))
                for cut in cuts:
                    row = export_cut_candidate(cut, bbox, sheet_name)
                    row["workbook_id"] = workbook_id
                    all_cut_rows.append(row)

    # 4. Write final_candidates.jsonl
    _write_jsonl(os.path.join(output_dir, "final_candidates.jsonl"), all_final_rows)

    # 5. Write features.csv
    _write_csv(os.path.join(output_dir, "features.csv"), all_training_rows)

    # 6. Write qc_events.jsonl
    _write_jsonl(os.path.join(output_dir, "qc_events.jsonl"), all_qc_rows)

    # 7. Write cut_scores.jsonl
    _write_jsonl(os.path.join(output_dir, "cut_scores.jsonl"), all_cut_rows)

    # 8. Write run_meta.json
    meta = {
        "workbook_id": workbook_id,
        "workbook_path": os.path.abspath(file_path),
        "timestamp": timestamp,
        "config_snapshot": cfg.__dict__,
        "sheet_count": len(workbook_results),
        "candidate_count": total_candidates,
    }
    with open(os.path.join(output_dir, "run_meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2, default=str)

    logger.info(
        "Run exported to %s (%d candidates across %d sheets)",
        output_dir, total_candidates, len(workbook_results),
    )
    return output_dir


def _write_jsonl(path: str, rows: List[dict]) -> None:
    """Write a list of dicts as newline-delimited JSON."""
    with open(path, "w", encoding="utf-8") as f:
        for row in rows:
            f.write(json.dumps(row, default=str) + "\n")


def _write_csv(path: str, rows: List[dict]) -> None:
    """Write a list of dicts as CSV with header row."""
    if not rows:
        with open(path, "w", encoding="utf-8") as f:
            pass  # empty file
        return
    fieldnames = list(rows[0].keys())
    with open(path, "w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
