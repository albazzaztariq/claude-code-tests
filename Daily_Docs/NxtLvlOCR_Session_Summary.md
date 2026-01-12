# NxtLvlOCR: Font-Aware Table OCR Training Session

**Date:** 2026-01-11
**Goal:** Train custom Tesseract models on actual table data to improve OCR accuracy
**Status:** Training data generated, ready to train model

---

## Problem Statement

Standard OCR engines (Tesseract, PaddleOCR) are trained on thousands of fonts. For academic papers using only 2-3 fonts (Times New Roman, Arial), this creates unnecessary ambiguity. Custom font-specific Tesseract models should improve accuracy.

---

## Session Progress

### Phase 1: Font Analysis (COMPLETED)
- **Task:** Analyze table images to identify fonts
- **Method:** Visual inspection + edge complexity analysis
- **Results:**
  - 11/13 tables use **serif fonts** (Times New Roman / Computer Modern) - LaTeX tables
  - 2/13 tables use **sans-serif fonts** (Arial / Helvetica)
- **Files:** Font classifications hardcoded in `font_aware_ocr.py`

### Phase 2: Ground Truth Preparation (COMPLETED)
- **Source:** Excel file `Anthropic_26Tbl_1-5-26.xlsx` (extracted by Claude API previously)
- **Location:** `MultiToolBenchmark/VLM_Test_Scripts/Results/`
- **Contains:** 16 sheets, each sheet = one table with structured data (cells in Excel cells)

**Tables excluded from training:**
- table_8_page9_2 (incorrect extraction)
- table_12_page2_1 (user rejected)

**Final training set:** 14 tables copied to `NxtLvlOCR/Crops/`:
1. table_3_page6_1
2. table_4_page4_1
3. table_4_page6_1
4. table_7_page3_1
5. table_7_page3_2
6. table_8_page5_1
7. table_9_page3_1
8. table_9_page6_1
9. table_10_page5_1
10. table_10_page5_2
11. table_10_page9_1
12. table_11_page3_1
13. table_11_page5_1
14. table_12_page2_2

### Phase 3: Preprocessing Pipeline Development (COMPLETED)

**Initial attempt:** Train directly on full table images → FAILED
**Reason:** Tesseract requires single-line images, not complex table layouts

**Solution:** Develop preprocessing to clean table images BEFORE OCR

#### Preprocessing Steps (in order):

1. **Grayscale Conversion**
   - RGB → Single channel (0-255 brightness values)
   - `cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)`

2. **Background Normalization** ← CRITICAL for gray backgrounds
   - **Problem:** Table 3_page6_1 had gray background (pixel value ~200)
   - **Solution:** Pixels > 200 → force to 255 (white)
   - Code: `result[result > 200] = 255`
   - **Why:** Without this, thresholding inverts gray backgrounds to black

3. **Sharpening (serif fonts) OR Denoising (sans-serif)**
   - **Serif:** Sharpen to enhance fine details (serifs)
     - Kernel: `[[-1,-1,-1], [-1,9,-1], [-1,-1,-1]]`
   - **Sans-serif:** Denoise to remove speckles
     - `cv2.fastNlMeansDenoising()`

4. **Contrast Enhancement**
   - Histogram equalization: `cv2.equalizeHist()`
   - Spreads pixel values across full 0-255 range
   - Makes text darker, background lighter

5. **Binary Thresholding**
   - Convert 256 grayscale values → 2 values (0=black, 255=white)
   - Otsu's automatic threshold: `cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)`
   - Result: Pure black text on pure white background

6. **Table Line Removal** ← CUSTOM ALGORITHM
   - **Problem:** Table grid lines interfere with OCR
   - **User's algorithm:** "If 10+ consecutive black pixels in straight row/col → it's a line"

   **Evolution of threshold:**
   - Initial: 10px → removed character parts
   - Tried: 20px → still too aggressive
   - Tried: 40px → better but missed some lines
   - Tried: 80px → removed title rows with hyphens
   - **Problem with title rows:** "Table 3 — Analysis" has long text stretches

   **Measured actual data:**
   - Longest word in titles: **28 pixels**
   - Table grid lines: **2700+ pixels**

   **Final threshold: 100 pixels**
   - Preserves all text (including title rows with hyphens)
   - Removes table grid lines

   **Final algorithm:**
   - Check for 1-3 consecutive rows/columns
   - Each row/column must have 100+ continuous black pixels
   - If all match → remove those rows/columns

**Final preprocessing script:** `full_pipeline_with_bg_norm.py`

**Output:** 14 preprocessed images in `preprocessed_final/` folder
**Quality:** User approved - "slightly grainy but let's try with this"

### Phase 4: Training Data Generation (COMPLETED)

**Approach:** Train on simple single-line images (not full tables)

**Step 1: Vocabulary Extraction**
- Script: `extract_vocabulary.py`
- Method: Parse all cells in ground_truth.xlsx, extract individual tokens
- **Result:** 597 unique tokens
- Examples: "Property", "PC1", "0.52", "%", "±", "Thickness"
- Saved to: `vocabulary.txt`

**Step 2: Training Image Generation**
- Script: `generate_training_from_vocab.py`
- Method:
  1. For each token, create 3 images (font sizes: 24, 28, 32pt)
  2. Use Times New Roman (serif font)
  3. Generate simple white background, black text image
  4. Save as TIF file
  5. Create ground truth file (.gt.txt) with exact text
  6. Generate box file (.box) using Tesseract makebox
- **Result:** 1791 training images generated
- Saved to: `training_images/` folder
- Training file list: `academic_tables.training_files.txt`

### Phase 5: Model Training (READY - NOT YET EXECUTED)

**Script to run:** `train_real_model.py` (already exists, previously failed but now has correct input)

**Training process:**
1. Read list of 1791 prepared files
2. Generate .tr files (Tesseract feature extraction)
3. Create unicharset (character set) - **batched to avoid Windows command-line length limit**
4. Create shape table
5. Cluster features
6. Combine into `academic_tables.traineddata` file
7. Install to Tesseract tessdata directory

**Previous failure:** Tried to train on full table images → box files didn't match
**Current status:** Now has proper single-line images → should work

---

## File Locations

```
NxtLvlOCR/
├── Crops/                          # 14 original table images
├── preprocessed_final/             # 14 preprocessed images (clean, lines removed)
├── training_images/                # 1791 training images (.tif, .box, .gt.txt)
│   └── academic_tables.training_files.txt
├── vocabulary.txt                  # 597 unique tokens
├── ground_truth.xlsx               # User-created (copied from Anthropic results)
│
├── full_pipeline_with_bg_norm.py   # Final preprocessing pipeline
├── extract_vocabulary.py           # Vocabulary extraction
├── generate_training_from_vocab.py # Training image generation
├── train_real_model.py             # Model training script (READY TO RUN)
└── font_aware_ocr.py               # OCR inference script
```

---

## Key Technical Decisions

### 1. Why Tesseract (not PaddleOCR, EasyOCR, etc.)?
**Only OCR engine that supports font-specific training.**
Other engines are deep learning models that can't be constrained to specific fonts at runtime without massive retraining.

### 2. Why train on simple images, not full tables?
Tesseract training requires:
- Single-line text images
- Character-level bounding boxes (.box files)
- Simple layouts for feature extraction

Training on complex table layouts fails because auto-generated box files don't match pixel locations.

### 3. Why preprocessing is critical?
Training happens on **simple clean images** (white background, black text, no lines).
Real tables have lines, shading, gray backgrounds.
**Preprocessing bridges the gap** - converts messy tables to look like training images.

### 4. Why 100px threshold for line removal?
- Longest word in title rows: 28px
- Table grid lines: 2700+px
- 100px is safely above words, safely below grid lines

---

## Next Steps (For Next Session)

1. **Train the model:**
   ```bash
   cd NxtLvlOCR
   python train_real_model.py
   ```
   - Should take 5-15 minutes
   - Output: `training_images/academic_tables.traineddata`

2. **Test on first table:**
   ```bash
   python font_aware_ocr.py --image "Crops/7_page3_table_1.png"
   ```
   - Should use custom model: `academic_tables`
   - Compare extracted text against ground_truth.xlsx

3. **If accuracy is good:**
   - Process all 14 tables
   - Calculate accuracy vs ground truth
   - Decide whether to proceed with full pipeline

4. **If accuracy is poor:**
   - Investigate: preprocessing issues? training data quality? model config?
   - May need to adjust preprocessing or add more training samples

---

## Common Issues & Solutions

### Issue: "FileNotFoundError: filename or extension too long"
**Cause:** Windows command-line limit (8191 chars)
**Solution:** Batch processing in train_real_model.py (processes 50 files at a time)

### Issue: Binary image has inverted colors (white text on black background)
**Cause:** Gray background wasn't normalized before thresholding
**Solution:** Add background normalization step (pixels > 200 → 255)

### Issue: Title rows being removed during line removal
**Cause:** Threshold too low, detecting words as "lines"
**Solution:** Measure actual word widths, set threshold above them (100px)

### Issue: Table grid lines not being removed
**Cause:** Threshold too high
**Solution:** Measure actual line widths, ensure threshold is below them

---

## Important Concepts

**Grayscale:** 256 shades (0=black, 255=white)
**Binary:** 2 values only (0=black, 255=white)
**Thresholding:** Converting grayscale → binary by picking cutoff value
**Otsu's method:** Automatic threshold selection algorithm
**Box file:** Character-level bounding boxes for Tesseract training
**Traineddata:** Tesseract's compiled model file format
**Token:** Individual word, number, or symbol (e.g., "Property", "0.52", "%")

---

## Font Classifications

### Serif (11 tables):
3_page6_1, 4_page6_1, 7_page3_1, 7_page3_2, 8_page5_1,
9_page3_1, 9_page6_1, 10_page5_1, 10_page5_2, 11_page3_1,
11_page5_1, 12_page2_2

### Sans-serif (2 tables):
4_page4_1, 10_page9_1

### Rotated:
4_page6_1 (90° counterclockwise)

---

## Session Notes

- User was VERY insistent on doing things correctly (no shortcuts)
- Preprocessing went through many iterations to get threshold right
- User provided feedback by viewing images ("still has lines", "graininess ok")
- Final preprocessing produces "slightly grainy" images but acceptable for OCR
- Training data generation completed successfully (1791 images)
- Ready to train model - next session should start with `python train_real_model.py`

---

**END OF SESSION SUMMARY**
