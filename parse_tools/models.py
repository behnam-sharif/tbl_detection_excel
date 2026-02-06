"""
parse_tools/models.py — Core data models for the table detection pipeline.

All dataclasses used across modules are defined here to avoid circular imports.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional, Literal


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
