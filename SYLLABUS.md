# xl_parse Course Syllabus

## Practical Excel Parsing for Health Economists: From Spreadsheets to Structured Data with Python + Claude

---

## A. Course Overview

### Audience

Health economists, outcomes researchers, and HTA modelers who:

- Work with Excel-based economic models daily (cost-effectiveness models, budget impact models, indirect treatment comparisons).
- Have little or no Python experience but are comfortable with Excel formulas and workbook navigation.
- Want to automate the tedious task of extracting tables from complex, multi-sheet workbooks.

### Prerequisites

- Comfort navigating Excel workbooks (sheets, ranges, merged cells, named ranges).
- A working computer (Windows 10/11 or macOS 12+) with admin rights to install software.
- A GitHub account (free tier is fine).
- Access to Claude (via Cursor IDE or claude.ai).
- No prior Python, Git, or command-line experience required.

### Learning Outcomes

By the end of this course, students will be able to:

1. Set up a Python development environment from scratch (virtual environment, dependencies, IDE).
2. Use Cursor IDE + Claude to write, debug, and iterate on Python code.
3. Use Git and GitHub for version control, collaboration, and backup.
4. Load Excel workbooks with openpyxl and inspect their contents programmatically.
5. Understand and run the xl_parse pipeline: from raw spreadsheet to detected table boundaries.
6. Read and modify heuristic thresholds to improve table detection on their own models.
7. Export structured table candidates for downstream ML, LLM, or manual review workflows.
8. Explain the mental model behind Repos 2 and 3 (labeling and knowledge-graph/RAG), even if they have not built them yet.

### Teaching Style

- **Excel-first analogies**: Every concept is introduced by relating it to something the student already knows in Excel.
- **Claude-assisted coding**: Students are encouraged to ask Claude for help at every step. The course teaches *how* to prompt effectively, not just how to type code.
- **Small, visible wins**: Each session ends with a tangible artifact (a screenshot, a JSON file, a Streamlit page) that proves something worked.
- **No memorization**: Students keep a personal cheat-sheet of commands and patterns. The goal is understanding, not recall.

### Tools Used

| Tool | Purpose |
|---|---|
| **Python 3.11+** | Programming language |
| **openpyxl** | Read Excel files from Python |
| **Streamlit** | Build interactive web UIs for the pipeline |
| **Cursor IDE** | Code editor with built-in Claude AI assistant |
| **Git + GitHub** | Version control and collaboration |
| **Claude (via Cursor)** | AI pair-programmer for writing and debugging code |

---

## D. Repo 1 Mental Model

Before diving into sessions, here is the big picture. The xl_parse pipeline processes **one sheet at a time** and produces a list of detected tables (bounding boxes + metadata).

```
                          ┌──────────────────────────────────────────────────┐
                          │              xl_parse  PIPELINE                  │
                          │           (runs per sheet, in parallel)          │
                          └──────────────────────────────────────────────────┘

  ┌─────────────┐    ┌─────────────────┐    ┌─────────────────────────────┐
  │  Excel File  │───▶│  SheetAccessor   │───▶│  Presence Matrix (0/1 grid) │
  │ (.xlsx)      │    │  (openpyxl wrap) │    │  1 = has content            │
  └─────────────┘    └─────────────────┘    │  0 = empty cell             │
                                             └──────────┬──────────────────┘
                                                        │
                                                        ▼
                                          ┌──────────────────────────┐
                                          │  Block Generation         │
                                          │  (connected components    │
                                          │   + gap bridging          │
                                          │   + header/label expand)  │
                                          └──────────┬───────────────┘
                                                     │
                                                     ▼
                                          ┌──────────────────────────┐
                                          │  Table Framing Heuristics │
                                          │  keep / reject / shrink   │
                                          │  (density, borders, etc.) │
                                          └──────────┬───────────────┘
                                                     │
                                                     ▼
                                          ┌──────────────────────────┐
                                          │  Grammar Heuristics       │
                                          │  (split multi-tables:     │
                                          │   separators, header      │
                                          │   restarts, changepoints) │
                                          └──────────┬───────────────┘
                                                     │
                                                     ▼
                                          ┌──────────────────────────┐
                                          │  QC Gate + Final          │
                                          │  Candidates               │
                                          │  (scores, header info,    │
                                          │   structural type, flags) │
                                          └──────────┬───────────────┘
                                                     │
                                                     ▼
                                          ┌──────────────────────────┐
                                          │  Export / Streamlit UI    │
                                          │  (JSON, ML rows, or      │
                                          │   interactive web view)   │
                                          └──────────────────────────┘
```

**Analogy for health economists**: Think of this like a factory line.

1. You feed in a workbook (your cost-effectiveness model).
2. The system reads each sheet (Inputs, Costs, Utilities, Results...).
3. For each sheet, it figures out which rectangular regions are actual tables versus labels, notes, or white space.
4. It scores each candidate table and gives you structured output: "rows 3-15, columns B-F is a cost table with headers in row 3."

---

## B. Syllabus Structure

### Session 1 — Environment Setup, First Code, and Git Foundations

*(Detailed below in Section C)*

**Learning Objectives:**

- Install Python, create a virtual environment, and install project dependencies.
- Launch Streamlit and understand the xl_parse project structure.
- Use Cursor + Claude to write and iterate on code.
- Initialize a Git repository, create a `.gitignore`, and make a first commit.
- Load an Excel workbook with openpyxl and print its sheet names.
- Write a minimal "parse per sheet" loop.

---

### Session 2 — The Presence Matrix: Turning a Sheet into a Grid of 0s and 1s

**Learning Objectives:**

- Explain what a presence matrix is and why it matters for table detection.
- Use `SheetAccessor` to open a specific sheet and read cell values.
- Build a presence matrix from scratch (conceptually), then use the repo's `get_presence_matrix_from_accessor()`.
- Visualize the matrix in the terminal and in Streamlit.

**Lecture Topics:**

- Why "has content vs. empty" is the foundation of all table detection.
- Excel analogy: conditional formatting that highlights non-empty cells. That is exactly what the presence matrix is, but as numbers.
- How `SheetAccessor` wraps openpyxl: dual workbooks (formulas vs. cached values), thread safety.
- The `BBox` dataclass: how the repo represents rectangular regions (0-indexed, inclusive).
- `bbox_to_excel_range()`: converting between code coordinates and Excel coordinates (A1:D10).

**Hands-On Lab:**

1. Open a sample workbook with `SheetAccessor`.
2. Print the grid dimensions (`nrows`, `ncols`).
3. Call `get_presence_matrix_from_accessor()` and print the result.
4. Write a small function that prints the matrix as a text grid (`.` for 0, `#` for 1).
5. Identify a known table region visually in Excel, then confirm it matches the `#` cluster.

**Common Mistakes & Debugging:**

- Forgetting that openpyxl is 1-indexed but `BBox` is 0-indexed.
- Loading a `.xls` file (openpyxl only handles `.xlsx`). Use the `xlrd` library or re-save as `.xlsx`.
- Sheets with merged cells can look "empty" in some cells; `SheetAccessor` handles this.

**Homework:**

- Pick any Excel model you have access to (or use the sample file). For the first three sheets, generate and save the presence matrix as a text file (one line per row, space-separated 0s and 1s).
- Write a 2-sentence summary of what you see: "Sheet 'Costs' has one dense block in the upper-left and scattered cells below."

**Stretch Goals:**

- Colorize the presence matrix in Streamlit using `st.dataframe()` with a custom styler.
- Count how many non-empty cells each sheet has and display a bar chart.

---

### Session 3 — Block Generation: Finding Connected Regions

**Learning Objectives:**

- Explain what "connected components" means in plain language.
- Understand gap bridging and why it is needed (small empty rows/columns between parts of the same table).
- Understand context expansion (adding header rows above and label columns to the left).
- Run `generate_block_candidates()` and inspect the output.

**Lecture Topics:**

- Connected components: "flood fill" from a non-empty cell, collecting all reachable non-empty neighbors (4-connectivity vs. 8-connectivity).
- Excel analogy: imagine selecting a non-empty cell and pressing Ctrl+Shift+End, but smarter — it stops at empty regions.
- Gap bridging: "If there's just one empty row between two filled regions, they're probably the same table." Controlled by `Config.bridging_rows` and `Config.bridging_cols`.
- Context expansion: a table often has a header row *above* the dense numeric region. The algorithm looks upward for text-heavy rows and leftward for label columns.
- `BlockCandidate` dataclass: `bbox`, `n_cells_on`, `density`, `provenance`, `sheet_name`.

**Hands-On Lab:**

1. Run `generate_block_candidates()` on a sheet with 2-3 obvious tables.
2. Print each candidate's Excel range and density.
3. Change `bridging_rows` from 2 to 0 and see how the blocks change (one big block splits into several).
4. Manually verify one block against the actual Excel sheet.

**Common Mistakes & Debugging:**

- Expecting one block per table — before framing, you might get one giant block that covers the whole sheet. That is normal; framing and splitting come later.
- Setting `min_on_cells` too high and filtering out small but real tables.
- Confusing `density` (fraction of non-empty cells within the bounding box) with "number of cells."

**Homework:**

- For a chosen sheet, run block generation with default config and with `bridging_rows=0`. Save both outputs. Write a short paragraph comparing them.

**Stretch Goals:**

- Overlay the bounding boxes on the presence matrix visualization from Session 2.
- Experiment with 8-connectivity (`Config.connectivity=8`) and describe what changes.

---

### Session 4 — Cell Features and Block Features: What Is In Each Cell?

**Learning Objectives:**

- Classify a cell's type: blank, text, number, or formula.
- Extract cell-level formatting: bold, fill color, borders, merged status, alignment.
- Aggregate cell features into block-level features: density profiles, type histograms, bold%, fill%, border coverage.
- Understand why these features matter for deciding "is this a real table?"

**Lecture Topics:**

- Cell classification priority: formula > blank > number > text.
- Excel analogy: you already know the difference between a cell that says "Total" (text), one that says "1,250.00" (number), and one that says "=SUM(B2:B10)" (formula). The classifier just formalizes this.
- Style extraction via openpyxl: `Font.bold`, `PatternFill`, `Border`, `Alignment`.
- `CellFeatures` dataclass: one per cell, captures type + formatting.
- `BlockFeatures` dataclass: one per block, aggregates across all cells in the block.
  - `row_density` / `col_density`: how full each row/column is.
  - `type_histogram`: count of blank/text/number/formula cells.
  - `bold_pct`, `fill_pct`, `border_coverage`: formatting statistics.
  - `internal_gap_rows` / `internal_gap_cols`: empty rows/columns inside the block.

**Hands-On Lab:**

1. Pick a single cell in a sample workbook. Use `classify_cell_type()` and `extract_cell_features()` on it. Print the result.
2. Use `classify_block_cells()` on one block to get the full cell-feature grid.
3. Use `compute_block_features()` to get the aggregated block features. Print the density profile and type histogram.
4. Compare two blocks: one that is a real cost table vs. one that is a title/notes region. See how the features differ.

**Common Mistakes & Debugging:**

- Assuming "number" means the cell *looks* like a number. Cells formatted as text but containing digits are classified as text.
- Forgetting that bold/fill are properties of the *cell style*, not the *value*. A bold cell with no value is still bold.
- `border_coverage` can be misleading on sheets where the modeler applied borders decoratively.

**Homework:**

- For one sheet, pick the largest block candidate. Export its `BlockFeatures` to a JSON file. Annotate (in a text file) which features would help you decide "this is a real table" vs. "this is a notes section."

**Stretch Goals:**

- Build a small Streamlit widget that lets you click a cell and see its `CellFeatures`.
- Create a summary table (in Excel or as a DataFrame) showing block features side by side for all blocks on one sheet.

---

### Session 5 — Table Framing Heuristics: Keep, Reject, or Shrink

**Learning Objectives:**

- Understand the three-outcome decision: keep, reject, or shrink.
- Trace how `evaluate_block()` uses density, borders, and headers to decide.
- Understand shrinking (trimming blank edges) and splitting (breaking on internal gaps).
- Run `run_framing_pipeline()` and interpret `TableDecision` objects.

**Lecture Topics:**

- The framing question: "Is this block-shaped region a real table, junk, or a table with extra whitespace?"
- Decision signals:
  - **Density**: tables are usually 50%+ filled. A region that is 10% filled is probably not a table.
  - **Borders**: if cells have drawn borders forming a box, it is likely a table even if density is low.
  - **Headers**: a text-heavy top row with bold or fill suggests a table header.
  - **Size**: a 1x1 block is not a table (controlled by `min_rows`, `min_cols`).
- Shrinking: if the block has blank rows at the bottom or blank columns on the right, trim them.
- Splitting: if there is a wide internal gap (empty rows/columns in the middle), try splitting into two blocks.
- Multi-pass refinement: after shrinking, re-evaluate. A shrunk block might now pass the density threshold.
- `TableDecision` dataclass: `action`, `rationale`, `confidence`, `signal_scores`.
- The `Config` thresholds that control all of this.

**Hands-On Lab:**

1. Run `run_framing_pipeline()` on all blocks from a sample sheet.
2. Print each block's action and rationale.
3. Find a "reject" decision. Read the rationale. Open the Excel sheet and verify — is the pipeline correct?
4. Find a "shrink" decision. Compare `original_bbox` vs. `final_bbox`. Check in Excel.
5. Change `min_density_keep` from 0.50 to 0.30 and re-run. See which blocks flip from reject to keep.

**Common Mistakes & Debugging:**

- Changing thresholds without understanding what they affect. Always change one threshold at a time.
- Expecting perfect results. The heuristics are a V1 baseline — they will make mistakes. That is why Repo 2 (labeling) and Repo 3 (ML/LLM) exist.
- Forgetting that framing runs *before* grammar splitting. A "keep" block might still be split later.

**Homework:**

- Run the framing pipeline on two different sheets. For each, create a small table (in Excel or Markdown) with columns: Block #, Excel Range, Action, Rationale (first 50 chars), Density. Highlight any decisions you disagree with.

**Stretch Goals:**

- Write a function that filters the pipeline output to only "keep" and "shrink" blocks and prints them in a clean summary.
- Experiment with `border_box_threshold` on a workbook that uses heavy border formatting.

---

### Session 6 — Grammar Heuristics: Splitting Multi-Table Blocks

**Learning Objectives:**

- Explain why a single block might contain multiple tables.
- Describe the three splitting strategies: separator detection, header restart detection, and changepoint detection.
- Understand how cut candidates are scored and selected.
- Run `run_grammar_splitting()` and trace a split decision.

**Lecture Topics:**

- The multi-table problem: in many HE models, the "Inputs" sheet has a cost table, a utility table, and a demographics table stacked vertically with just one or two blank rows between them. Block generation might merge them into one block.
- Strategy 1 — Separator detection: look for nearly-empty rows (or columns) that span most of the block width. These are the "breathing room" between stacked tables.
- Strategy 2 — Header restart: look for rows that resemble a header (text-heavy, bold, filled) appearing *after* a block of numbers. This signals "a new table just started."
- Strategy 3 — Changepoint detection: look for a sharp shift in the text-vs-number ratio from one row to the next. If rows 3-12 are 80% numbers and row 13 is 90% text, something changed.
- `CutCandidate` dataclass: each potential split boundary gets scores from each strategy plus a `combined_score`.
- Cut selection and application: the top-scoring candidates are selected (up to `max_splits_per_block`), then the block is sliced at those rows/columns.
- Post-split validation: sub-blocks that are too small are discarded.

**Hands-On Lab:**

1. Find a sheet where block generation produces one large block that clearly contains 2+ tables.
2. Run `run_grammar_splitting()` and inspect the `CutCandidate` list.
3. For each cut candidate, print its index, axis, and per-strategy scores.
4. Verify: does the selected cut row/column match the visual boundary in Excel?
5. Adjust `hdr_bold_ratio` and see if the header-restart strategy finds different cuts.

**Common Mistakes & Debugging:**

- Confusing row-axis cuts (horizontal splits) with column-axis cuts (vertical splits).
- Expecting grammar splitting to fix blocks that framing already rejected — splitting only runs on kept/shrunk blocks.
- Over-splitting: if thresholds are too aggressive, you get many tiny fragments. The QC gate (next session) catches some of these.

**Homework:**

- For a multi-table sheet, document the grammar splitting output: how many cuts were proposed, which were selected, and what the resulting sub-blocks are. Include the Excel ranges.

**Stretch Goals:**

- Visualize the cut candidates on a presence matrix: draw a red line at each proposed cut row.
- Compare results with `max_splits_per_block=1` vs. `max_splits_per_block=5`.

---

### Session 7 — QC Gate, Final Candidates, and Scoring

**Learning Objectives:**

- Explain the purpose of the post-split QC gate (catch bad splits, merge tiny fragments, refine headers/footers).
- Read and interpret `FinalCandidate` objects: bbox, header info, structural type, scores, QC flags.
- Understand structural types: input, calc, mixed, lookup.
- Run the full pipeline end-to-end via `parse_workbook()`.

**Lecture Topics:**

- QC gate actions: pass, reject, merge, header pullback, footer pullback, re-split.
- Header pullback: if the top row(s) of a detected table are actually a title or banner (one merged cell spanning the width), pull them out of the table.
- Footer pullback: if the bottom row(s) are a "Source:" note, remove them.
- Tiny-fragment merge: if two adjacent sub-blocks are both very small, merge them back.
- `FinalCandidate` dataclass:
  - `header`: which rows are headers, confidence, multi-row flag.
  - `key_columns`: which column is the label/stub column.
  - `structural_type`: "input" (mostly text/numbers entered by hand), "calc" (mostly formulas), "mixed", or "lookup" (narrow text-keyed table).
  - `table_likeness_score`, `header_score`, `confidence`.
  - `qc_flags`: any warnings (e.g., "low_density", "banner_header_removed").
- `SheetResult`: the container for all per-sheet output.

**Hands-On Lab:**

1. Run `parse_workbook()` on a full workbook (all sheets).
2. For each sheet, print the number of final candidates.
3. Pick one `FinalCandidate` and walk through every field. Map `bbox` back to the Excel range.
4. Check the `structural_type` — does it match what you see in Excel (input table vs. formula-heavy results sheet)?
5. View the results in the Streamlit app (`main.py`). Compare the UI view to the raw Python output.

**Common Mistakes & Debugging:**

- Running `parse_workbook()` on a huge file (50+ sheets) and wondering why it is slow. Start with 2-3 sheets using the `sheet_names` parameter.
- Ignoring `qc_flags` — they are warnings, not errors, but they tell you where the pipeline is uncertain.
- Confusing `table_likeness_score` (how much this looks like *any* table) with `confidence` (how confident the pipeline is in its specific boundaries).

**Homework:**

- Run the full pipeline on a workbook of your choice. Export all `FinalCandidate` objects to a JSON file (use Python's `dataclasses.asdict()`). Write a one-paragraph summary: "The pipeline found X tables across Y sheets. The highest-confidence table is [range] on sheet [name] (confidence=Z). The lowest-confidence table is..."

**Stretch Goals:**

- Build a small summary Streamlit page that shows a table of all final candidates across all sheets, sorted by confidence.
- Compare `parse_workbook()` output with `max_workers=1` vs. `max_workers=4` and time both.

---

### Session 8 — Export, Config Tuning, and Real-World Debugging

**Learning Objectives:**

- Use `ml_export.py` to export training rows and cut candidates.
- Understand the LLM QC payload stub and what it is designed for (Repo 3 preview).
- Tune `Config` thresholds for a specific workbook and document the changes.
- Debug a case where the pipeline gets a table boundary wrong.

**Lecture Topics:**

- `export_training_row()`: flattening a block + decision + features into one dict for ML.
- `export_cut_candidate()`: flattening a cut for analysis.
- `build_llm_qc_payload()`: what it would send to an LLM for quality review (stub in V1).
- Config tuning strategy:
  1. Run pipeline with defaults.
  2. Identify the worst failure (missed table or false positive).
  3. Hypothesize which threshold is responsible.
  4. Change it, re-run, compare.
  5. Document the change and why.
- Real-world debugging patterns:
  - "The pipeline missed a table" → check presence matrix, then check block generation, then check framing.
  - "The pipeline merged two tables" → check gap bridging, then check grammar splitting.
  - "The boundary is off by a few rows" → check context expansion, then check shrinking.

**Hands-On Lab:**

1. Run the full pipeline and export training data for one sheet.
2. Open the exported JSON. Identify the fields that would be useful for training a classifier.
3. Intentionally pick a failure case. Trace it through the pipeline: presence matrix → blocks → framing → grammar → QC.
4. Find the stage where the error was introduced. Adjust the relevant config parameter. Re-run and verify the fix.
5. Save your tuned `Config` as a JSON file for reproducibility.

**Common Mistakes & Debugging:**

- Trying to tune too many thresholds at once. Change one, test, then change the next.
- Fixing one sheet and breaking another. Always test on at least 2-3 sheets after a config change.
- Forgetting to document *why* you changed a threshold. Future-you will thank present-you.

**Homework:**

- Choose a workbook where the pipeline makes at least one mistake. Document the mistake, your diagnosis, the config change you made, and the result (before/after). This is your "debugging case study."

**Stretch Goals:**

- Write a script that runs the pipeline on 3 workbooks with the same config and produces a summary report.
- Experiment with `build_llm_qc_payload()` — print the payload and imagine what an LLM would say about it.

---

### Session 9 — Putting It All Together: End-to-End Workflow and Streamlit App

**Learning Objectives:**

- Run the full xl_parse workflow from file upload to exported results.
- Customize the Streamlit app for your own needs.
- Use Git to manage your changes professionally (branches, commits, push).
- Present your results to a colleague (non-technical communication of technical output).

**Lecture Topics:**

- The Streamlit app walkthrough: upload, sheet viewer, block candidates, table detection.
- Customizing the app: adding your own visualization, changing the layout, adding download buttons.
- Git workflow for real projects: feature branches, meaningful commit messages, pull requests.
- Communicating results: "The pipeline detected 12 tables across 5 sheets. Here are the ones it got right and the ones we need to review."

**Hands-On Lab:**

1. Run the Streamlit app. Upload a workbook. Walk through every tab.
2. Add a "Download Results as JSON" button to the Streamlit app.
3. Create a Git branch for your changes. Commit. Push to GitHub.
4. Prepare a 3-slide summary of your pipeline results for a non-technical colleague.

**Common Mistakes & Debugging:**

- Editing `main.py` while Streamlit is running and getting confused by hot-reload behavior.
- Committing data files (`.xlsx`) to Git. Add `*.xlsx` to `.gitignore`.
- Writing commit messages like "update" or "fix". Be specific: "Add JSON download button to Streamlit results page."

**Homework:**

- Push your customized Streamlit app to GitHub. Share the repo link.
- Write a short README section describing what your customization does.

**Stretch Goals:**

- Add a "Compare Two Workbooks" feature to the Streamlit app.
- Set up GitHub Actions to run a basic smoke test on push.

---

### Session 10 — What Comes Next: Preview of Repo 2 and Repo 3

**Learning Objectives:**

- Explain the limitations of V1 heuristics and why ML/LLM approaches are needed.
- Describe the role of Repo 2 (table labeling and annotation).
- Describe the role of Repo 3 (knowledge graph / RAG for structured extraction).
- Identify which parts of Repo 1 feed into Repos 2 and 3.

**Lecture Topics:**

- Where V1 heuristics break down:
  - Tables with no clear header row.
  - Tables embedded in free-text notes.
  - Sheets with complex merged-cell layouts.
  - Inconsistent formatting across modelers.
- Repo 2 preview — Labeling:
  - Annotating detected tables with ground truth (correct/incorrect boundaries, table type, header rows).
  - Building a training dataset from your own workbooks.
  - Why human review is still essential (the "human-in-the-loop" principle).
- Repo 3 preview — Knowledge Graph + RAG:
  - Extracting *meaning* from tables, not just boundaries (e.g., "this is a cost table with drug names in column A and per-cycle costs in column D").
  - Building a knowledge graph that connects tables across sheets and workbooks.
  - Using RAG (retrieval-augmented generation) to answer questions like "What is the utility value for stable disease in the base case?"
- How Repo 1 output feeds downstream:
  - `FinalCandidate` → Repo 2 (the candidates to be labeled).
  - `export_training_row()` → Repo 2 (features for ML training).
  - `build_llm_qc_payload()` → Repo 3 (the stub that becomes a real LLM call).

**Hands-On Lab:**

1. Run the full pipeline and export `FinalCandidate` objects.
2. For 3 candidates, manually annotate: "correct boundary? Y/N. If N, what should it be?" This is exactly what Repo 2 will automate.
3. For 1 candidate, write a natural-language description of the table: "This is a 5x3 cost table with headers [Drug, Annual Cost, Source] and 4 data rows." This is exactly what Repo 3 will generate.
4. Discuss: what would you need from an LLM to do step 3 automatically?

**Common Mistakes & Debugging:**

- Thinking Repos 2/3 replace Repo 1. They build on it. Better V1 heuristics mean less work for the ML/LLM layers.
- Rushing to label data before understanding the pipeline output. Always run and inspect Repo 1 first.

**Homework:**

- Write a one-page reflection: What did you learn in this course? Which part of the pipeline is most useful for your work? What would you want Repo 2 or Repo 3 to do?

**Stretch Goals:**

- Sketch a data model for the labeling system (Repo 2). What fields would you need?
- Read about RAG (retrieval-augmented generation) and write three questions you would want to ask a knowledge graph built from your workbooks.

---

## C. Session 1 — Detailed Plan

### Title: Environment Setup, First Code, and Git Foundations

### Duration: ~2.5 hours (lecture + lab)

### Session Overview

This is the foundation session. By the end, every student will have a working Python environment, a running Streamlit app, a Git repository, and their first Python script that loads an Excel file. No prior experience required — we start from zero.

---

### Topic 1: Python Virtual Environment Setup

**Why this matters:** A virtual environment (venv) is like a separate "project folder" for Python packages. It keeps this project's libraries separate from anything else on your computer. Think of it like having a dedicated USB drive for each project — nothing leaks between them.

**What you will do:**

1. Open a terminal (PowerShell on Windows, Terminal on macOS).
2. Verify Python is installed.
3. Create a virtual environment.
4. Activate it.
5. Install the project dependencies.

**Windows (PowerShell):**

```powershell
# Check Python version (must be 3.11 or higher)
python --version

# Navigate to your project folder
cd C:\Users\YourName\Desktop\work\xl_parse

# Create a virtual environment named "venv"
python -m venv venv

# Activate the virtual environment
.\venv\Scripts\Activate.ps1

# Your prompt should now show (venv) at the beginning
# Install project dependencies
pip install -r requirements.txt
```

**macOS (Terminal):**

```bash
# Check Python version (must be 3.11 or higher)
python3 --version

# Navigate to your project folder
cd ~/Desktop/work/xl_parse

# Create a virtual environment named "venv"
python3 -m venv venv

# Activate the virtual environment
source venv/bin/activate

# Your prompt should now show (venv) at the beginning
# Install project dependencies
pip install -r requirements.txt
```

**How to verify it worked:**

```bash
# This should show openpyxl and streamlit in the list
pip list
```

**Common mistake:** Forgetting to activate the venv before installing packages. If you see packages installing "globally," close the terminal, reopen it, navigate to the project folder, and activate again.

---

### Topic 2: Project Structure and Streamlit

**Why this matters:** Before writing any code, you need to understand what files exist in the project and what they do. Then you will launch the Streamlit app to see the pipeline in action.

**Project folder structure:**

```
xl_parse/
├── .gitignore              ← tells Git which files to ignore
├── requirements.txt        ← list of Python packages needed
├── pipeline.py             ← the main pipeline orchestrator
├── main.py                 ← Streamlit web app (the UI)
├── block_generation.py     ← finds connected regions in the sheet
├── table_framing_heuristics.py  ← decides keep/reject/shrink
├── grammar_heuristics.py   ← splits multi-table blocks
├── post_split_qc_gate.py   ← quality control after splitting
├── final_candidates.py     ← builds final table candidates
├── ml_export.py            ← exports data for ML / LLM use
└── parse_tools/            ← shared utilities
    ├── __init__.py          ← package init, re-exports public APIs
    ├── models.py            ← all data structures (BBox, Config, etc.)
    ├── sheet_accessor.py    ← reads Excel files safely
    ├── presence.py          ← builds 0/1 presence matrices
    ├── cell_classifier.py   ← classifies cell type + style
    ├── feature_extraction.py← aggregates features per block
    └── id_utils.py          ← ID generation utilities
```

**Launch the Streamlit app:**

```bash
# Make sure your venv is activated!
streamlit run main.py
```

This opens a browser tab at `http://localhost:8501`. You should see the xl_parse web interface.

**What to look for:**

- An "Upload Excel" button.
- A sheet selector dropdown (after uploading).
- Block candidate and table detection views.

**To stop Streamlit:** Press `Ctrl+C` in the terminal.

---

### Topic 3: Cursor IDE + Claude Workflow

**Why this matters:** Cursor is a code editor with a built-in AI assistant (Claude). For non-technical users, Claude is your co-pilot: you describe what you want in plain English, and Claude writes or explains the code. This is the most important skill in the course.

**Opening the project in Cursor:**

1. Open Cursor IDE.
2. File → Open Folder → select `xl_parse`.
3. You should see the file tree on the left.

**How to talk to Claude in Cursor:**

- Press `Ctrl+L` (Windows) or `Cmd+L` (macOS) to open the Claude chat.
- Type your question or request in plain English.

**Effective prompts (good vs. bad):**

```
BAD:  "help me with python"
GOOD: "Show me how to open the file 'model.xlsx' using openpyxl and print the sheet names."

BAD:  "fix this"
GOOD: "I'm getting a FileNotFoundError on line 5. The file 'model.xlsx' is in the same
       folder as this script. What's wrong?"

BAD:  "write the whole pipeline"
GOOD: "Write a function that takes a file path and returns a dictionary where keys are
       sheet names and values are the number of non-empty cells in each sheet."
```

**The Claude iteration loop:**

1. **Ask** Claude to write or explain something.
2. **Read** the output. Do you understand it? If not, ask "explain line 3 in simpler terms."
3. **Run** the code. Did it work?
4. **If error:** paste the error message back to Claude and say "I got this error when running the code above. How do I fix it?"
5. **If works:** move on. Save the code.

**Try it now:**

```python
# Type this in a new file called scratch.py and ask Claude:
# "What does this code do, line by line?"

import openpyxl

wb = openpyxl.load_workbook("sample.xlsx")
print(wb.sheetnames)
wb.close()
```

---

### Topic 4: Git, GitHub, and .gitignore

**Why this matters:** Git tracks every change you make to your code. GitHub stores a backup online. If you break something, you can go back to a working version. Think of it like "Track Changes" in Word, but for code.

**One-time Git setup (if not already done):**

```bash
# Set your name and email (used in commit messages)
git config --global user.name "Your Name"
git config --global user.email "you@example.com"
```

**Understanding .gitignore:**

The `.gitignore` file tells Git which files to *not* track. You do not want to upload virtual environments, cached files, or large Excel data to GitHub.

The project already has a `.gitignore`:

```
# Virtual environments
venv/
venv1/
.venv/

# Python cache
__pycache__/
*.pyc

# OS junk
.DS_Store

# VSCode (optional)
.vscode/
```

**Recommended additions for your work:**

```
# Data files (don't commit patient data or large models!)
*.xlsx
*.xls
data/

# Cursor settings
.cursor/
```

**Basic Git workflow:**

```bash
# Check the current status (what has changed?)
git status

# Stage all changes for commit
git add .

# Commit with a descriptive message
git commit -m "Add scratch.py with basic openpyxl workbook loading"

# Push to GitHub (first time: set upstream)
git push -u origin main

# After first time, just:
git push
```

**The three states of a file in Git:**

1. **Modified** — you changed it but have not staged it.
2. **Staged** — you ran `git add` on it; it is ready to be committed.
3. **Committed** — it is saved in your local Git history.

**Golden rule:** Commit often, with clear messages. "Add presence matrix visualization for Costs sheet" is better than "stuff."

---

### Topic 5: First Code — Loading an Excel Workbook

**Why this matters:** Everything in xl_parse starts with loading a workbook. This is the equivalent of "opening the file" — but in Python, so you can automate what happens next.

**Create a file called `session1_explore.py`:**

```python
import openpyxl

# ---------------------------------------------------
# Step 1: Load a workbook and print sheet names
# ---------------------------------------------------
file_path = "sample.xlsx"  # <-- replace with your file

wb = openpyxl.load_workbook(file_path, data_only=True)
print(f"Workbook: {file_path}")
print(f"Number of sheets: {len(wb.sheetnames)}")
print(f"Sheet names: {wb.sheetnames}")
print()

# ---------------------------------------------------
# Step 2: For each sheet, print dimensions and a preview
# ---------------------------------------------------
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    print(f"--- {sheet_name} ---")
    print(f"  Rows: {ws.max_row}, Columns: {ws.max_column}")

    # Print the first 3 rows as a preview
    for row_idx, row in enumerate(ws.iter_rows(max_row=3, values_only=True), start=1):
        # Convert None to "" for cleaner display
        clean_row = [str(cell) if cell is not None else "" for cell in row]
        print(f"  Row {row_idx}: {clean_row[:6]}")  # first 6 columns
    print()

wb.close()
```

**Run it:**

```bash
python session1_explore.py
```

**Expected output (your numbers will differ):**

```
Workbook: sample.xlsx
Number of sheets: 5
Sheet names: ['Inputs', 'Costs', 'Utilities', 'Model', 'Results']

--- Inputs ---
  Rows: 25, Columns: 8
  Row 1: ['Parameter', 'Value', 'Source', '', '', '']
  Row 2: ['Discount rate', '0.035', 'NICE ref case', '', '', '']
  Row 3: ['Time horizon (years)', '40', 'Base case', '', '', '']

--- Costs ---
  Rows: 18, Columns: 6
  ...
```

---

### Topic 6: Iterating Over Sheets — A Minimal "Parse per Sheet" Function

**Why this matters:** The xl_parse pipeline processes each sheet independently. This is how you will structure your own code: a function that handles one sheet, called in a loop.

**Add this to `session1_explore.py` (or create `session1_parse.py`):**

```python
from parse_tools import SheetAccessor, get_presence_matrix_from_accessor

def parse_single_sheet(file_path: str, sheet_name: str) -> dict:
    """
    Minimal parse function for one sheet.
    Returns a summary dict with basic info.
    """
    accessor = SheetAccessor(file_path, sheet_name)
    try:
        # Build the presence matrix
        presence = get_presence_matrix_from_accessor(accessor)

        # Count non-empty cells
        total_cells = accessor.nrows * accessor.ncols
        non_empty = sum(sum(row) for row in presence)

        return {
            "sheet_name": sheet_name,
            "rows": accessor.nrows,
            "cols": accessor.ncols,
            "total_cells": total_cells,
            "non_empty_cells": non_empty,
            "fill_rate": round(non_empty / total_cells, 3) if total_cells > 0 else 0,
        }
    finally:
        accessor.close()


# ---------------------------------------------------
# Run the parse function on every sheet
# ---------------------------------------------------
import openpyxl
import json

file_path = "sample.xlsx"  # <-- replace with your file

wb = openpyxl.load_workbook(file_path, read_only=True)
sheet_names = wb.sheetnames
wb.close()

results = []
for name in sheet_names:
    print(f"Parsing: {name}...")
    summary = parse_single_sheet(file_path, name)
    results.append(summary)
    print(f"  {summary['non_empty_cells']} / {summary['total_cells']} cells filled "
          f"({summary['fill_rate']:.1%})")

# Save results to JSON
with open("session1_output.json", "w") as f:
    json.dump(results, f, indent=2)

print(f"\nResults saved to session1_output.json")
```

**Run it:**

```bash
python session1_parse.py
```

**Expected output:**

```
Parsing: Inputs...
  87 / 200 cells filled (43.5%)
Parsing: Costs...
  54 / 108 cells filled (50.0%)
Parsing: Utilities...
  32 / 60 cells filled (53.3%)
Parsing: Model...
  312 / 1500 cells filled (20.8%)
Parsing: Results...
  45 / 120 cells filled (37.5%)

Results saved to session1_output.json
```

**Key pattern to remember:** This is the same structure the real pipeline uses. Look at `pipeline.py`:

- `process_single_sheet()` is the "real" version of our `parse_single_sheet()`.
- `parse_workbook()` is the "real" version of our loop — but it runs sheets in parallel using `ThreadPoolExecutor`.

---

### Session 1 Student Tasks (Checklist)

Complete all of the following before the next session. Check each box as you go.

**Setup Tasks:**

- [ ] Python 3.11+ installed and verified (`python --version` shows 3.11 or higher).
- [ ] Virtual environment created and activated (prompt shows `(venv)`).
- [ ] `pip install -r requirements.txt` completed successfully.
- [ ] Cursor IDE installed and project folder opened.
- [ ] Git configured (`git config --global user.name` and `user.email` set).

**Code Tasks:**

- [ ] `session1_explore.py` runs and prints sheet names + previews.
- [ ] `session1_parse.py` runs and prints fill rates for all sheets.
- [ ] `session1_output.json` exists and contains valid JSON.
- [ ] Streamlit app launches (`streamlit run main.py` → browser opens).

**Git Tasks:**

- [ ] `.gitignore` updated with `*.xlsx` and `data/` entries.
- [ ] First commit made with a descriptive message.
- [ ] Code pushed to GitHub (or local Git repo created if no GitHub access yet).

**Claude Tasks:**

- [ ] Asked Claude at least one question in Cursor and got a useful answer.
- [ ] Pasted an error message to Claude and got a fix.

### Expected Artifacts to Submit / Screenshot

| Artifact | Format | What it proves |
|---|---|---|
| Terminal showing `python --version` | Screenshot | Python is installed |
| Terminal showing `(venv)` prompt + `pip list` | Screenshot | Venv is active, packages installed |
| Terminal output of `session1_explore.py` | Screenshot or text paste | Workbook loads, sheet names print |
| Terminal output of `session1_parse.py` | Screenshot or text paste | Parse loop works, fill rates print |
| `session1_output.json` | File | JSON export works |
| Browser showing Streamlit app | Screenshot | Streamlit runs |
| `git log --oneline` output | Screenshot or text paste | Git is working, commit exists |

---

## E. Practical HE Modeling Context

Throughout this course, examples use language familiar to health economists:

| Excel Term | xl_parse Term | Example |
|---|---|---|
| "Inputs sheet" | Sheet with low formula density, high text/number mix | Discount rates, time horizons, population sizes |
| "Costs sheet" | Sheet with cost tables stacked vertically | Drug costs, admin costs, AE costs |
| "Utilities sheet" | Sheet with utility/disutility values | EQ-5D values by health state |
| "Model sheet" | Sheet with high formula density | Markov trace, transition probabilities |
| "Results sheet" | Sheet with summary tables, often formatted | ICERs, NMB, tornado diagram inputs |
| "Table" | A `FinalCandidate` with a `BBox` | Any rectangular region with headers and data |
| "Merged header" | `is_merged=True` in `CellFeatures` | "Base Case" spanning columns B-F |
| "Bold row" | `is_bold=True`, high `bold_pct` | A header row or section title |

No proprietary or patient-level data is needed. Students can use any publicly available HE model (e.g., NICE technology appraisal models available on the NICE website) or create a simple sample workbook with fake data.

---

*End of Syllabus — xl_parse Repo 1 Course*
