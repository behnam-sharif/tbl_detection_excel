### What each file does

| File | Role |
|------|------|

| **parse_tools.py** | Builds the **presence matrix** (0/1 per cell) and **parses ranges** (optional for later use): given a file, sheet, and cell range, returns cell_matrix, formatting, formulas, and/or cell_data. Used by the loader to get bounding boxes and cell data for summaries. |
| **block_generation.py** | Takes a sheet’s presence matrix and produces **block candidates** (bboxes): finds connected components of non-empty cells, optionally bridges small gaps, expands to include headers/labels, and returns a list of `BlockCandidate` (bbox, cell count, density, provenance). |
| **main.py** | Streamlit app: upload Excel, parse (via loader or fallback), view sheets, and generate block candidates for all sheets (using block_generation). |

---

Per Sheet
  |
  v
[1] Load sheet grid (values + formulas + formatting)
  |
  v
[2] Build 0/1 Presence Matrix
    - 1 if value OR formula exists
    - 0 otherwise
  |
  v
[3] Connected-Components Block Generation   ✅ (added)
    - segment contiguous “1” regions into BlockCandidates
    - allow small-gap bridging (optional)
    - expand blocks to capture nearby header/title rows & label columns
  |
  v
[4] Cell-Level Classification (within each BlockCandidate)
    - type: blank/text/number/formula
    - style: bold/fill/borders/merge/alignment
    - compute row/col density + type histograms
  |
  v
[5] Table Framing Heuristics (per BlockCandidate)
    - “is this table-like?”
    - header-likelihood at top
    - border box / gridlines (if available)
    - column consistency / low internal gaps
    -> keep / reject / shrink / expand
  |
  v
[6] Grammar Heuristics (split multi-tables) 
    - detect header restarts
    - detect separators (blank rows/cols)
    - detect change points (text→num patterns)
    -> split into SubCandidates
  |
  v
[7] Post-Split QC Gate (fast filters)       ✅ (added)
    - min size (>=2x2, etc.)
    - blank ratio threshold
    - must-have header score OR stable structure
    - reject banner/title-only blocks
    - merge tiny fragments if needed
  |
  v
[8] Output Final Candidates (for ML/LLM later)
    - bbox (r1,c1,r2,c2)
    - header rows / key columns
    - structural type (input/calc/mixed/lookup)
    - scores + qc flags + provenance