# Config Folder - Configuration Files Reference

This folder contains all configuration files for the ScholarSweep pipeline.

---

## 📋 Files Overview

| File | Purpose | Used By |
|------|---------|---------|
| **MetricsSearchable.txt** | Master list of searchable metric terms and acronyms | Text extraction, Table OCR, Final dataset |
| **MetricsAcronymMapping.txt** | Reference mapping of acronyms to phrases (auto-generated) | Reference only |
| **query.txt** | OpenAlex API search query configuration | GenerateCorpus |
| **FilterPrompt.txt** | LLM system prompt for paper relevance filtering | LLM Filter |
| **LLMTextExtractPrompt.txt** | LLM system prompt for extracting metric values from sentences | LLM Text Extract |
| **JournalList.py** | Preferred journal list for ranking papers | GenerateCorpus |

---

## 🔑 MetricsSearchable.txt - Core Configuration

**Purpose**: Master list of all metric terms (phrases and acronyms) to search for in papers.

### Structure

The file contains **phrases** followed by their **acronyms**:

```
Wetting Time Top          ← Phrase (has lowercase letters)
WTT                       ← Acronym (all uppercase)
Wetting Time Bottom
WTB
```

### Rules for Adding Terms

#### 1. **Phrases come BEFORE acronyms**
   - Always list the full phrase first
   - Then list the acronym(s) on the next line(s)

#### 2. **Acronym Format**
   - All uppercase letters (A-Z)
   - Can include: numbers (0-9), hyphens (‑ or -), special chars (α, ″, ₜ, ᵦ)
   - NO lowercase letters

#### 3. **Hyphenated Acronyms**
   For hyphenated acronyms (e.g., `AR‑J`, `MC‑J`), the **base phrase must exist** before the hyphenated version:

   ```
   Absorption Rate           ← Base phrase (required)
   AR‑J                      ← Hyphenated acronym
   ```

#### 4. **Multiple Acronyms per Phrase**
   You can list multiple acronyms for the same phrase:

   ```
   Water Vapor Transmission Rate
   WVTR
   MVTR                      ← Alternative acronym for same phrase
   ```

#### 5. **Order Matters**
   The acronym mapping function looks **backwards** from each acronym to find the nearest phrase above it.

### How Acronym Mapping Works

The `load_acronym_mapping()` function in `NativeTextExtract_pipeline.py`:

1. Reads MetricsSearchable.txt line by line
2. When it finds an acronym (all uppercase), it looks **backwards** to find the first phrase (has lowercase)
3. Creates mapping: `{acronym: phrase}`

**Example**:
```
Line 22: Maximum Absorption Rate    ← Phrase
Line 23: MAR                         ← Acronym
Line 24: Maximum Absorption Rate Top ← Phrase
Line 25: MART                        ← Acronym
```

Creates mappings:
- `MAR -> Maximum Absorption Rate`
- `MART -> Maximum Absorption Rate Top`

### Adding New Metrics

**Step 1**: Add the full phrase
```
Surface Temperature Recovery Time
```

**Step 2**: Add the acronym on the next line
```
Surface Temperature Recovery Time
STR
```

**Step 3**: (Optional) Add hyphenated variants
```
Surface Temperature Recovery Time
STR
STR‑J
```

**Step 4**: Regenerate the mapping file (automatic at runtime, or run manually):
```bash
cd Config
python -c "
import re
from pathlib import Path

metrics_file = Path('MetricsSearchable.txt')
with open(metrics_file, 'r', encoding='utf-8') as f:
    lines = [line.strip() for line in f if line.strip()]

acronym_to_phrase = {}

for i, term in enumerate(lines):
    if re.match(r'^[A-Z0-9α‑\-″ₜᵦ\s\(\)/]+$', term) and not term.replace(' ', '').replace('-', '').replace('‑', '').isdigit():
        for j in range(i-1, -1, -1):
            prev_term = lines[j]
            if re.search(r'[a-z]', prev_term):
                acronym_to_phrase[term] = prev_term
                break

output = []
output.append('# Acronym to Phrase Mapping')
output.append('# Generated from MetricsSearchable.txt')
output.append(f'# Total: {len(acronym_to_phrase)} acronym mappings')
output.append('')
output.append('=== ACRONYM MAPPINGS ===')

for acr, phrase in sorted(acronym_to_phrase.items()):
    output.append(f'{acr} -> {phrase}')

with open('MetricsAcronymMapping.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(output))

print(f'Regenerated with {len(acronym_to_phrase)} mappings')
"
```

---

## 📄 MetricsAcronymMapping.txt - Reference File

**Purpose**: Human-readable reference showing all acronym-to-phrase mappings.

**How it's created**: Auto-generated from MetricsSearchable.txt using the script above.

**Format**:
```
AR‑J -> Absorption Rate
MAR -> Maximum Absorption Rate
MART -> Maximum Absorption Rate Top
MARB -> Maximum Absorption Rate Bottom
MC -> Moisture Content
MVTR -> Water Vapor Transmission Rate
```

**Note**: This file is for **reference only**. The code generates mappings dynamically from MetricsSearchable.txt at runtime.

---

## 🔍 How Metric Matching Works in Filtered Text

When papers are processed, the filtered text files show:

**For Phrases**:
```
Metric Match: Wetting Time
The fabric showed a wetting time of 5.2 seconds.
```

**For Acronyms**:
```
Metric Match: WTT, which means Wetting Time Top
The WTT was measured at 3.1 seconds for the test sample.
```

This format makes it clear what each acronym stands for when reviewing filtered text.

---

## 🔧 query.txt - OpenAlex Search Configuration

**Purpose**: Defines the search query for OpenAlex API to find relevant papers.

**Full Structure**:
```
# ScholarSweep Config File - Test 10 PDFs
# Moisture management in textiles (woven/knit)

[APIs]
2

[Max Results]
10

[Journals]


[OpenAlex]
field: fulltext
terms: woven///knit /AND/ moisture-wicking///moisture wicking
```

**Section Breakdown**:

- **Comments** (lines starting with `#`): Description of the query
- **[APIs]**: Number of OpenAlex API keys to use (1-2)
- **[Max Results]**: Maximum number of papers to download (e.g., 10, 50, 100)
- **[Journals]**: Leave blank for no journal filtering
- **[OpenAlex]**:
  - `field:` Search field (`fulltext`, `title`, `abstract`)
  - `terms:` Search terms with operators:
    - `///` = OR operator (e.g., `woven///knit` = "woven OR knit")
    - `/AND/` = AND operator (joins term groups)
    - Example: `term1///term2 /AND/ term3///term4` = "(term1 OR term2) AND (term3 OR term4)"

**How to create from scratch**:
1. Copy the structure above
2. Modify the comment to describe your search
3. Set number of APIs (usually 2 for parallel downloads)
4. Set max results (start with 10 for testing)
5. Leave [Journals] blank unless you have a specific journal filter
6. Set search field (recommend `fulltext`)
7. Define search terms using `///` for OR and `/AND/` for AND

**Example for battery research**:
```
# Battery cathode materials research

[APIs]
2

[Max Results]
50

[Journals]


[OpenAlex]
field: fulltext
terms: lithium///sodium /AND/ cathode///anode /AND/ capacity///performance
```

---

## 💬 FilterPrompt.txt - LLM Filter System Prompt

**Purpose**: System prompt for the SambaNova LLM to filter papers for relevance.

**Used by**: `3_LLMFilter/llm_filter.py`

**Full Structure**:
```
You are a research paper evaluator.

Follow these steps EXACTLY:

1. Does this paper perform an experiment?
   - If NO: Respond with ONLY the text "No experiment" and STOP.
   - If YES: Continue to step 2.

2. Does the experiment involve [YOUR CRITERIA HERE]?
   - If NO: Respond with ONLY the text "No [criteria]" and STOP.
   - If YES: Respond with ONLY the text "Passes" and STOP.

IMPORTANT:
- Your response must be EXACTLY one of: "No experiment", "No [criteria]", or "Passes"
- Do NOT add explanations, reasoning, or additional text
- Do NOT use quotes around your response

Paper text:
{text}
```

**Key Components**:

1. **Step 1 - Experiment Check**: Always check if the paper performs experiments (filters out review papers, theoretical papers)
2. **Step 2 - Domain Check**: Replace `[YOUR CRITERIA HERE]` with your domain-specific requirements
3. **Response Format**: Must return EXACTLY one of three strings:
   - `"No experiment"` - Paper doesn't perform experiments
   - `"No [criteria]"` - Paper doesn't meet domain criteria
   - `"Passes"` - Paper meets all criteria
4. **{text} placeholder**: Required - LLM filter script replaces this with actual paper text

**How to create from scratch**:
1. Copy the structure above
2. Replace `[YOUR CRITERIA HERE]` in Step 2 with your domain requirements
3. Update the response strings to match your criteria
4. Keep the `{text}` placeholder at the end

**Example for textile research**:
```
You are a research paper evaluator.

Follow these steps EXACTLY:

1. Does this paper perform an experiment?
   - If NO: Respond with ONLY the text "No experiment" and STOP.
   - If YES: Continue to step 2.

2. Does the experiment involve knit fabrics, woven fabrics, or yarns?
   - If NO: Respond with ONLY the text "No knit/woven/yarn" and STOP.
   - If YES: Respond with ONLY the text "Passes" and STOP.

IMPORTANT:
- Your response must be EXACTLY one of: "No experiment", "No knit/woven/yarn", or "Passes"
- Do NOT add explanations, reasoning, or additional text
- Do NOT use quotes around your response

Paper text:
{text}
```

**Example for battery research**:
```
You are a research paper evaluator.

Follow these steps EXACTLY:

1. Does this paper perform an experiment?
   - If NO: Respond with ONLY the text "No experiment" and STOP.
   - If YES: Continue to step 2.

2. Does the experiment involve battery electrodes or electrolytes?
   - If NO: Respond with ONLY the text "No battery materials" and STOP.
   - If YES: Respond with ONLY the text "Passes" and STOP.

IMPORTANT:
- Your response must be EXACTLY one of: "No experiment", "No battery materials", or "Passes"
- Do NOT add explanations, reasoning, or additional text
- Do NOT use quotes around your response

Paper text:
{text}
```

---

## 📊 LLMTextExtractPrompt.txt - Metric Value Extraction

**Purpose**: System prompt for the SambaNova LLM to extract numeric metric values from filtered sentences.

**Used by**: `4a_LLMTextExtract/llm_extract_values.py`

**Current Prompt**:
```
This first line of each group of text contains the a metric keyword or phrase that was found in sentences along with numbers.  Return the metric value if stated in the sentence.  For example, if the sentence is, "The OMMC value was measured.  Fabric 1's value was 0.6", then the value is 0.6.  If the sentence is "WTT was 2.3s.", then the value is 2.3.  But if the sentence is Wetting Time Top was measured for Fabric 1", or "The absorption rate for 1 of 3 fabrics was large.  Table 2 lists statistics from the experiment", then there is no corresponding value listed for those metrics and you should return "No value found".
```

**How it works**:
1. LLM receives filtered text in format:
   ```
   Metric Match: <metric name>
   <sentence with potential value>
   ```
2. LLM analyzes sentence for numeric values
3. LLM returns:
   - The numeric value (e.g., "0.6", "2.3")
   - "No value found" if no clear value exists

**How to modify**:
1. Edit the prompt text to adjust extraction logic
2. Test with sample sentences: `python llm_extract_values.py`
3. Adjust examples and rules as needed

**Example Input/Output**:

| Input Sentence | Expected Output |
|---------------|----------------|
| "The OMMC value was measured. Fabric 1's value was 0.6" | 0.6 |
| "WTT was 2.3s." | 2.3 |
| "Wetting Time Top was measured for Fabric 1" | No value found |
| "The absorption rate for 1 of 3 fabrics was large." | No value found |

**Optimization TODO**:
- Current prompt: ~180 words
- Target: <100 words (reduce API costs)
- Strategy: Remove redundant examples, use concise phrasing

---

## 📚 JournalList.py - Preferred Journals

**Purpose**: Dictionary of high-quality journals for ranking downloaded papers.

**Used by**: GenerateCorpus when selecting which papers to download first.

**Full Structure**:
```python
"""
JournalList.py - [Domain] Journal List for ScholarSweep

Contains [N] journal entries with full names and acronyms.
Used for journal filtering in API searches.
"""

# =============================================================================
# [DOMAIN] JOURNAL LIST ([N] entries)
# =============================================================================
# Format: "Full Name": "ACRONYM"
TEXTILE_JOURNALS = {
    "Textile Research Journal": "TRJ",
    "Journal of Applied Polymer Science": "JAPS",
    "Fibers and Polymers": "FP",
    # ... more journals
}

# =============================================================================
# REVERSE LOOKUP: ACRONYM -> LIST OF FULL NAMES
# =============================================================================
ACRONYM_TO_JOURNALS = {}
for name, acronym in TEXTILE_JOURNALS.items():
    if acronym:
        if acronym not in ACRONYM_TO_JOURNALS:
            ACRONYM_TO_JOURNALS[acronym] = []
        ACRONYM_TO_JOURNALS[acronym].append(name)
```

**Key Components**:

1. **TEXTILE_JOURNALS dict**: Maps full journal names to acronyms
   - Format: `"Full Journal Name": "ACRONYM"`
   - Empty string `""` for journals without acronyms
2. **ACRONYM_TO_JOURNALS dict**: Reverse lookup (auto-generated)
   - Maps acronym to list of full names
   - Handles multiple journals with same acronym

**How to create from scratch**:

1. **Create the file structure**:
```python
"""
JournalList.py - Battery Research Journal List for ScholarSweep

Contains 50 journal entries with full names and acronyms.
Used for journal filtering in API searches.
"""

# Format: "Full Name": "ACRONYM"
BATTERY_JOURNALS = {
    "Journal of Power Sources": "JPS",
    "Energy Storage Materials": "ESM",
    "ACS Energy Letters": "ACSEL",
    "Advanced Energy Materials": "AEM",
    "Journal of the Electrochemical Society": "JES",
    # Add more journals...
}

# Reverse lookup (copy this exactly)
ACRONYM_TO_JOURNALS = {}
for name, acronym in BATTERY_JOURNALS.items():
    if acronym:
        if acronym not in ACRONYM_TO_JOURNALS:
            ACRONYM_TO_JOURNALS[acronym] = []
        ACRONYM_TO_JOURNALS[acronym].append(name)
```

2. **Rename the dictionary** to match your domain:
   - `TEXTILE_JOURNALS` → `BATTERY_JOURNALS`, `DRUG_JOURNALS`, etc.
   - Update the `for` loop to use the new name

3. **Add journals**:
   - One entry per line: `"Full Name": "ACRONYM",`
   - Use empty string `""` if no acronym
   - Common acronym sources: journal website, Google Scholar, Web of Science

4. **Find relevant journals** for your domain:
   - Search Google Scholar for papers in your field
   - Check journal rankings (Impact Factor, SJR)
   - Look at references in key papers
   - Use journal databases (Scopus, Web of Science)

---

## 🗑️ Temporary/Analysis Files

These files are created during development and can be safely deleted:

- `missing_base_acronyms.txt` - Analysis output for finding missing base phrases
- `substring_conflicts.txt` - Analysis output for substring matching conflicts
- `MetricsSearchable_backup.txt` - Backup copy of MetricsSearchable.txt
- `nul:` - Windows null redirect artifact (should be deleted by cleanup daemon)

---

## 🔄 Workflow: Adding a New Domain

If you want to adapt ScholarSweep for a different research domain (e.g., battery materials, drug discovery):

### Step 1: Replace MetricsSearchable.txt
Create a new list of metrics relevant to your domain:
```
Battery Capacity
Specific Capacity
Cycle Life
Coulombic Efficiency
CE
```

### Step 2: Create query.txt
See the **query.txt section** above for complete creation instructions. Example:
```
# Battery cathode materials research

[APIs]
2

[Max Results]
50

[Journals]


[OpenAlex]
field: fulltext
terms: lithium///sodium /AND/ cathode///anode /AND/ capacity///performance
```

### Step 3: Create FilterPrompt.txt
See the **FilterPrompt.txt section** above for complete creation instructions. Example:
```
You are a research paper evaluator.

Follow these steps EXACTLY:

1. Does this paper perform an experiment?
   - If NO: Respond with ONLY the text "No experiment" and STOP.
   - If YES: Continue to step 2.

2. Does the experiment involve battery electrodes or electrolytes?
   - If NO: Respond with ONLY the text "No battery materials" and STOP.
   - If YES: Respond with ONLY the text "Passes" and STOP.

IMPORTANT:
- Your response must be EXACTLY one of: "No experiment", "No battery materials", or "Passes"
- Do NOT add explanations, reasoning, or additional text
- Do NOT use quotes around your response

Paper text:
{text}
```

### Step 4: (Optional) Create JournalList.py
See the **JournalList.py section** above for complete creation instructions. Example:
```python
BATTERY_JOURNALS = {
    "Journal of Power Sources": "JPS",
    "Energy Storage Materials": "ESM",
    "ACS Energy Letters": "ACSEL",
    # ...
}

ACRONYM_TO_JOURNALS = {}
for name, acronym in BATTERY_JOURNALS.items():
    if acronym:
        if acronym not in ACRONYM_TO_JOURNALS:
            ACRONYM_TO_JOURNALS[acronym] = []
        ACRONYM_TO_JOURNALS[acronym].append(name)
```

### Step 5: Run the pipeline
```bash
python Main.py
```

The rest of the pipeline (text extraction, filtering, table OCR) works the same way!

---

## 📊 Current Metrics Statistics

**MetricsSearchable.txt**:
- Total entries: 229
- Phrases: 135
- Acronyms: 94
- Unique metric concepts: ~100+

**Coverage**: Moisture management, thermal properties, mechanical properties, permeability, wicking, absorption, drying

---

## 🛠️ Troubleshooting

### Problem: Acronym not mapping correctly
**Solution**: Check that the base phrase exists BEFORE the acronym in MetricsSearchable.txt

**Example (WRONG)**:
```
AR‑J                          ← Acronym BEFORE phrase
Absorption Rate               ← Phrase AFTER acronym
```

**Example (CORRECT)**:
```
Absorption Rate               ← Phrase BEFORE acronym
AR‑J                          ← Acronym AFTER phrase
```

### Problem: MetricsAcronymMapping.txt is outdated
**Solution**: Regenerate it using the script in the "Adding New Metrics" section above

### Problem: Too many false positives in filtered text
**Solution**:
1. Check for substring conflicts in MetricsSearchable.txt
2. Use more specific phrases instead of generic terms
3. Adjust the sentence filtering logic in `extract_metric_sentences()`

---

## 📝 Notes

- All files use UTF-8 encoding (required for special characters like α, ‑, ₜ, ᵦ)
- MetricsSearchable.txt is case-insensitive for matching (regex uses `re.IGNORECASE`)
- Word boundary matching is used (`\b` regex) to avoid false matches inside words
- The pipeline allows multiple matches per sentence (same sentence can match multiple metrics)
