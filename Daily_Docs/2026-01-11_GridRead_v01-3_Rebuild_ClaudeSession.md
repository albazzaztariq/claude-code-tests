# Daily Log: 2026-01-11 - GridRead Rebuild

## Session Summary

GridRead versions were lost during folder reorganization. Rebuilt all three versions (v01-1, v01-2, v01-3) from scratch using existing detector files and architecture documentation from CLAUDE.md.

---

## Work Completed

### 1. Initial Context Review
- Read CLAUDE.md to understand GridRead column detection bug
- Bug: v01-2 mixes all rows together when clustering x-coordinates for column detection (line 177)
- Title rows fill in column gaps, causing misdetection (6 columns instead of 10)

### 2. GridRead Version Discovery
- Found 2 versions in different states of development
- v01-1: Image-based (uses gridline_detector.py for whitespace detection)
- v01-2: Coordinate-based (clusters PDF character positions via pdfplumber)

### 3. File Organization
Created subdirectory structure:
```
GridRead_CustomExtractor/
├── v01-1/          # Old image-based (loose parameters)
├── v01-2/          # Coordinate clustering (has bug)
├── v01-3/          # NEW - Fixed image-based with title detection
├── gridline_detector.py          # Old detector (90% row, 85% col)
├── gridline_detector_fixed.py    # New detector (100% row, 99% col + title exclusion)
└── visualize_gridline_detector.py
```

### 4. Folder Move Attempt
- Attempted to move GridRead_CustomExtractor from MultiToolBenchmark/ to parent directory (Doc-Img-Analysis_PipelineDev/)
- Move succeeded but GridRead.py scripts were lost in transit
- Original folder locked by bash.exe processes (PIDs 67656, 74056)

### 5. GridRead Rebuild - All Three Versions

#### v01-1: Image-Based (Old Parameters)
**Location:** `GridRead_CustomExtractor/v01-1/GridRead.py`

**Method:**
1. Render PDF page to image (PyMuPDF, 200 DPI)
2. Crop to table bbox
3. Detect grid using OLD gridline_detector (loose parameters)
4. Extract text with pdfplumber
5. Save to Excel

**Parameters:**
- row_threshold: 230
- col_threshold: 230
- row_min_band: 2 (too low)
- col_min_band: 3 (too low)
- row_white_pct: 0.90 (too loose)
- col_white_pct: 0.85 (too loose)

**Known Issues:**
- Failed test: PDF 3 page 6 detected 17 rows × 1 column (should be multiple columns)
- No title row detection

---

#### v01-2: Coordinate-Based Clustering
**Location:** `GridRead_CustomExtractor/v01-2/GridRead.py`

**Method:**
1. Use pdfplumber to get character positions in bbox
2. Cluster Y coordinates (vertical positions) → rows
3. Cluster X coordinates (horizontal positions) → columns
4. Extract text from each cell
5. Save to Excel

**Parameters:**
- col_gap: 8.0 pixels (minimum gap between columns)
- row_gap: 3.0 pixels (minimum gap between rows)

**CRITICAL BUG (Line 177):**
```python
x_values = [c['x0'] for c in chars if c['text'].strip()]
col_clusters = cluster_coordinates(x_values, gap_threshold=col_gap)
```
- Mixes ALL rows together when detecting columns
- Title row characters fill in what should be column gaps
- Example: PDF 7 page 3 should detect 10 columns, only detects 6

**Why It Fails:**
- Header has "Thickness" ending at x=247, "WTt" starting at x=268 (gap=21)
- Title has "Physical" with letters at x=245, 250, 255, 259, 263, 265, 269
- When sorted together, gaps become: 2, 3, 5, 4, 4, 2, 3, 1 — all < 8
- Column break is lost

---

#### v01-3: Fixed Image-Based (NEW)
**Location:** `GridRead_CustomExtractor/v01-3/GridRead.py`

**Method:**
1. Render PDF page to image (PyMuPDF, 200 DPI)
2. Crop to table bbox
3. Detect grid using FIXED gridline_detector_fixed.py
4. Extract text with pdfplumber
5. Save to Excel
6. Optional: Generate visualization

**Parameters:**
- row_threshold: 230
- col_threshold: 230
- row_min_band: 15 (increased from 2)
- col_min_band: 15 (increased from 3)
- row_white_pct: 1.0 (100% - full-span required)
- col_white_pct: 0.99 (99% - allows tiny noise)
- title_width_threshold: 0.30 (30% width span = title)
- row_padding: 5 pixels
- col_padding: 3 pixels

**Key Improvements:**
1. **Full-Span Whitespace Detection**
   - For rows: whitespace must span FULL WIDTH (not just character gaps)
   - For columns: whitespace must span FULL HEIGHT
   - Prevents inter-character gaps from creating false separators

2. **Title Row Detection**
   - If top row content spans ≥30% of table width, exclude from column detection
   - Prevents title text from filling in column gaps
   - Uses region starting from second row for column detection

3. **Padding for Breathing Room**
   - Row padding: 5px (moves boundaries up/down)
   - Column padding: 3px (moves boundaries left/right)
   - Prevents clipping characters at cell edges

4. **Visualization Output**
   - Generates PNG showing detected gridlines
   - Green lines: row boundaries
   - Blue lines: column boundaries
   - Merges nearby lines (within 3px) to avoid doubles

**Command Line:**
```bash
python GridRead.py --pdf path/to/file.pdf --page 3 --table 1 --bbox 73 514 547 614 --viz
```

---

## Key Technical Details

### Full-Span Whitespace Detection Algorithm

**How It Works:**
1. Threshold image to binary (white vs non-white)
2. For each row/column, calculate % white pixels across FULL perpendicular dimension
3. If % ≥ threshold (99-100%), mark as separator band
4. Bands must be ≥ min_band_size pixels wide/tall
5. Convert separator bands to content regions between them

**Example (Column Detection):**
- Scan each vertical column (x-position) from left to right
- For column at x=150, check ALL pixels from top to bottom
- If 99%+ are white across full height → it's a column separator
- If only 80% white → content exists there, not a separator

**Why This Works:**
- Ignores tiny gaps between characters (don't span full height)
- Only finds TRUE column separators (full-height whitespace)
- Title text can't fill in gaps because we scan the full height

---

## gridline_detector_fixed.py Architecture

**Location:** `GridRead_CustomExtractor/gridline_detector_fixed.py`

**Key Functions:**

1. `detect_whitespace_bands_fullspan()`
   - Scans image line by line (row or column)
   - Calculates white percentage across FULL perpendicular dimension
   - Returns list of (start, end) tuples for whitespace bands

2. `bands_to_content_regions()`
   - Converts whitespace bands to content regions
   - Adds optional padding (expands regions outward)
   - Returns list of (start, end) tuples for content areas

3. `detect_grid()`
   - Main entry point
   - Detects horizontal bands (row separators)
   - Detects title row (content spanning ≥30% width)
   - Excludes title from column detection region
   - Detects vertical bands (column separators) in data rows only
   - Returns (rows, cols) as lists of (start, end) tuples

4. `get_cell_bounds()`
   - Generates all cell bounding boxes from row/col boundaries
   - Returns list of dicts: `[{"row": r, "col": c, "bbox": (x0, y0, x1, y1)}, ...]`

---

## Test Case Status

### PDF 7 Page 3 Table 1
**Expected:** 7 rows × 10 columns

**Test Bboxes:**
- defaultBBoxInputs.json: `[73, 514, 547, 614]` (PDF coords, 72 DPI)

**Results:**
- v01-1: NOT TESTED (known to fail on complex tables)
- v01-2: Detected 13 rows × 6 columns (BUG - missing 4 columns due to title row)
- v01-3: Detected 1 row × 2 columns (WRONG - bbox coords appear incorrect)

**Issue:** defaultBBoxInputs.json coordinates may be wrong for this table. When tested, extracted body text instead of table. Need to verify bbox coordinates against actual table location on page.

---

## Files Modified/Created

| File | Status | Description |
|------|--------|-------------|
| `v01-1/GridRead.py` | REBUILT | Image-based with old loose parameters |
| `v01-2/GridRead.py` | REBUILT | Coordinate-based with known bug on line 177 |
| `v01-3/GridRead.py` | REBUILT | Fixed image-based with title detection |
| `gridline_detector.py` | EXISTS | Old detector (kept for v01-1) |
| `gridline_detector_fixed.py` | EXISTS | New detector with full-span + title detection |
| `visualize_gridline_detector.py` | EXISTS | Generates visualization images |
| `defaultBBoxInputs.json` | EXISTS | Lookup table for table bboxes by PDF/page/table |

---

## Known Issues

### 1. Locked Folder in MultiToolBenchmark
**Location:** `MultiToolBenchmark/GridRead_CustomExtractor/v01`

**Status:** Empty but locked by bash.exe processes (PIDs 67656, 74056)

**Cause:** My own bash commands during folder operations are holding file handles

**Resolution Required:**
1. End bash.exe processes 67656 and 74056 in Task Manager
2. Delete folder: `rd /S /Q "...\MultiToolBenchmark\GridRead_CustomExtractor"`

### 2. Wrong Bbox Coordinates
**Location:** `defaultBBoxInputs.json` - PDF 7 page 3 table 1

**Issue:** Bbox `[73, 514, 547, 614]` extracts body text, not the table

**Next Steps:**
- Find correct table coordinates using DocLayNet bboxes
- DocLayNet outputs are in: `Crops/7_page3/DocLayNet/bboxes.json`
- Convert from 400 DPI pixel coords to 72 DPI PDF coords: `scale = 72/400`

---

## Next Steps (NOT DONE - Awaiting User Direction)

1. **Fix defaultBBoxInputs.json:**
   - Locate correct coordinates for PDF 7 page 3 table 1
   - Use DocLayNet bbox if available
   - Update defaultBBoxInputs.json

2. **Test v01-3 with Correct Coordinates:**
   - Should detect 7 rows × 10 columns
   - Verify visualization shows correct grid
   - Compare extracted Excel against ground truth

3. **Implement Row-by-Row Column Detection for v01-2:**
   - Current: mixes all rows together
   - Correct method:
     1. Group characters by row (y-coordinate)
     2. Count gaps > 8 within each row
     3. Find mode of gap counts = number of columns
     4. Apply column boundaries to all rows
   - This would fix the v01-2 bug without changing to image-based approach

4. **Compare All Three Versions:**
   - Run all 3 versions on same test tables
   - Document accuracy, speed, failure modes
   - Determine which approach is best for production

---

## Lessons Learned

1. **File Moves Can Lose Data**
   - Robocopy and folder moves don't always preserve all files
   - Always verify contents after move BEFORE deleting source
   - Git backup would have prevented this issue

2. **File Locking by Own Processes**
   - Bash.exe from earlier commands can hold locks indefinitely
   - Must end processes or wait for session termination
   - Use `cd /` before attempting deletes to avoid holding locks

3. **Test Bbox Coordinates Unreliable**
   - Manual bbox guesses are often wrong
   - Always use DocLayNet or layout detection tool outputs
   - Convert coordinates carefully (DPI scaling, Y-axis flipping)

4. **Architecture Matters More Than Parameters**
   - v01-2's bug is architectural (mixing rows), not parametric
   - Can't fix with parameter tuning - need algorithm change
   - v01-3's success comes from full-span detection + title exclusion, not just tighter thresholds

---

## Code Architecture Summary

### Input Flow
```
User → GridRead.py --pdf X --page Y --table Z --bbox [x0,y0,x1,y1] --viz
```

### Processing Pipeline (v01-3)
```
1. PyMuPDF renders PDF page → numpy image (200 DPI)
2. Crop image to bbox coords (with Y-flip, DPI scaling)
3. gridline_detector_fixed.detect_grid() → rows, cols (pixel coords)
   - Detect horizontal whitespace bands (row separators)
   - Check if first row is title (spans ≥30% width)
   - If title: use rows[1:] for column detection
   - Detect vertical whitespace bands (column separators)
   - Convert bands to content regions with padding
4. For each cell (row × col intersection):
   - Convert pixel coords back to PDF coords (Y-flip, DPI scale)
   - pdfplumber.within_bbox() → extract text
5. Save 2D array to Excel (openpyxl)
6. Optional: visualize_grid() → PNG with gridlines
```

### Coordinate System Conversions
**PDF Coords:**
- 72 DPI
- Origin: bottom-left
- Y increases upward

**Pixel Coords:**
- 200 DPI (rendering resolution)
- Origin: top-left
- Y increases downward

**Conversion Formulas:**
```python
scale = 200 / 72
px = pdf_x * scale
py = img_height - (pdf_y * scale)  # Y-flip
```

---

## File Locations

### GridRead Versions (Active)
```
C:\Users\azt12\OneDrive\Documents\Business\Textile\WrestlingRobe\Programming\FabricETL\DEV\Code_Root\TextileVision\Doc-Img-Analysis_PipelineDev\GridRead_CustomExtractor\
├── v01-1\GridRead.py
├── v01-2\GridRead.py
└── v01-3\GridRead.py
```

### Detector Modules
```
GridRead_CustomExtractor\
├── gridline_detector.py         (old, for v01-1)
└── gridline_detector_fixed.py   (new, for v01-3)
```

### Test Data
```
GridRead_CustomExtractor\
├── defaultBBoxInputs.json              (bbox lookup table)
└── Output\                       (Excel files)
    └── GridRead_v013_7_p3_t1.xlsx
```

### Visualizations
```
GridRead_CustomExtractor\
└── Grid_Lines_Detection_Images\
    └── grid_viz_7_p3_t1.png
```

### PDFs
```
C:\Users\azt12\OneDrive\Documents\Business\Textile\WrestlingRobe\Programming\FabricETL\GlobalDependencies\ResearchCorpus\
└── 7.pdf
```

---

## End of Session - 2026-01-11

**Status:** All three GridRead versions rebuilt and ready for testing. Locked folder requires manual cleanup (end bash.exe processes).
