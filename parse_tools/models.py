"""
parse_tools/models.py — Core data models for the table detection pipeline.

All dataclasses used across modules are defined here to avoid circular imports.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Any, List, Dict, Tuple, Optional, Literal


@dataclass
class BBox:
    """Axis-aligned bounding box, 0-indexed inclusive on all sides."""
    r1: int  # top row
    c1: int  # left col
    r2: int  # bottom row
    c2: int  # right col

    def clip(self, nrows: int, ncols: int) -> BBox:
        return BBox(
            max(0, self.r1), max(0, self.c1),
            min(nrows - 1, self.r2), min(ncols - 1, self.c2),
        )

    def area(self) -> int:
        return (self.r2 - self.r1 + 1) * (self.c2 - self.c1 + 1)

    def height(self) -> int:
        return self.r2 - self.r1 + 1

    def width(self) -> int:
        return self.c2 - self.c1 + 1

    def __iter__(self):
        return iter((self.r1, self.c1, self.r2, self.c2))


@dataclass
class BlockCandidate:
    """A rectangular region that *may* contain a table."""
    bbox: BBox
    n_cells_on: int
    density: float
    provenance: str       # e.g. "cc", "cc+bridge", "cc+expand"
    sheet_name: str = ""  # populated by generate_block_candidates


@dataclass
class CellFeatures:
    """Per-cell classification and style information."""
    cell_type: str  # "blank" | "text" | "number" | "formula"
    is_bold: bool = False
    has_fill: bool = False
    border_flags: Tuple[bool, bool, bool, bool] = (False, False, False, False)  # top, bottom, left, right
    is_merged: bool = False
    h_alignment: Optional[str] = None  # left/center/right/general
    v_alignment: Optional[str] = None  # top/center/bottom


@dataclass
class BlockFeatures:
    """Aggregated features for an entire block."""
    row_density: List[float] = field(default_factory=list)
    col_density: List[float] = field(default_factory=list)
    type_histogram: Dict[str, int] = field(default_factory=dict)
    row_type_histograms: List[Dict[str, int]] = field(default_factory=list)
    col_type_histograms: List[Dict[str, int]] = field(default_factory=list)
    internal_gap_rows: List[int] = field(default_factory=list)
    internal_gap_cols: List[int] = field(default_factory=list)
    bold_pct: float = 0.0
    fill_pct: float = 0.0
    border_coverage: float = 0.0
    style_histogram: Dict[str, float] = field(default_factory=dict)


@dataclass
class TableDecision:
    """Result of heuristic evaluation for a single block."""
    action: str       # "keep" | "reject" | "shrink" | "expand"
    rationale: str
    original_bbox: BBox
    final_bbox: Optional[BBox] = None  # None if reject
    confidence: float = 0.5
    signal_scores: Dict[str, float] = field(default_factory=dict)
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class Config:
    """Centralized thresholds & weights — no magic numbers elsewhere."""
    # Block generation
    bridging_rows: int = 2
    bridging_cols: int = 1
    connectivity: int = 4
    min_on_cells: int = 4
    do_bridging: bool = True

    # Framing heuristics
    min_density_keep: float = 0.50
    min_density_reject: float = 0.15
    min_cells_reject: int = 4
    min_rows: int = 2
    min_cols: int = 2
    border_box_threshold: float = 0.60
    max_internal_gap_ratio: float = 0.30

    # Grammar — separator
    sep_max_density: float = 0.10
    sep_min_cross_span: float = 0.75

    # Grammar — header restart
    hdr_text_ratio: float = 0.60
    hdr_bold_ratio: float = 0.40
    hdr_fill_ratio: float = 0.30
    hdr_min_filled: float = 0.50
    body_numeric_ratio: float = 0.50
    body_min_rows_before: int = 3

    # Grammar — change-point
    cp_min_shift: float = 0.30
    cp_label_col_max: int = 2

    # Grammar — split constraints
    max_splits_per_block: int = 5
    min_sub_block_height: int = 2
    min_sub_block_width: int = 2
    min_sub_block_cells: int = 4

    # Pipeline
    max_framing_passes: int = 2
    max_workers: int = 4

    # Export flags (off by default)
    export_training_data: bool = False
    export_cut_candidates: bool = False


@dataclass
class CutCandidate:
    """A potential split line scored by each detection strategy."""
    index: int              # 0-based relative row/col index within block
    axis: str               # "row" | "col"
    separator_score: float = 0.0
    header_restart_score: float = 0.0
    changepoint_score: float = 0.0
    combined_score: float = 0.0
    selected: bool = False
