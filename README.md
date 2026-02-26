## PPT → Excel Test Sheet Converter

This is a small local tool that converts compliance-matrix style PPT/PPTX files into structured QA test Excel sheets.

### Feature overview

- **Backend** uses **FastAPI + python-pptx + openpyxl**:
  - Accepts uploaded `.pptx` (and `.ppt` if LibreOffice is available).
  - Scans slides that contain the sentence `This is the Compliance matrix that has been applied to`
    anywhere in a title placeholder or text box.
  - For each such occurrence, starts a **new test project** which becomes a **separate Excel tab**.
    The sheet name is derived from the text after that sentence (for example
    `the Contractor Class High-Risk`), until the next occurrence or end of the file.
  - Within each relevant slide/table:
    - Finds the `Flag` header column.
    - Each **row** under `Flag` is treated as a question with multiple conditions.
    - Question text comes from the `Flag` column.
    - Condition columns are `Approved`, `Approved with Restriction`, `Not Approved`;
      the tool looks at the cell text in those columns to build multiple Excel rows.
    - Yes/No style rows like:
      - `Approved` = `Has Answered Yes`
      - `Not Approved` = `Has Answered No or Question Unanswered`
      are expanded into three Excel rows with `Unanswered / No / Yes` in column C, and
      `Green/Red` pre-filled in column D according to the business rules.
    - Generic condition rows (e.g. KPI data states) are turned into one row per non-empty
      condition, with:
      - **Column B**: question text
      - **Column C**: condition description (text from the condition cell)
      - **Column D**: expected outcome color label `Green / Yellow / Red` based on the column.
    - Grey empty rows in the PPT table are converted into grey separator rows in Excel.
  - Notes handling:
    - If the slide or table contains
      `Please note: the following factors only apply if you have answered YES`,
      a gate row is inserted between the surrounding questions:
      **`If anwered Yes to above queestion`** in column B.
    - If it contains
      `Please note: the following factors only apply if you have answered NO`,
      a gate row **`If anwered No to above queestion`** is inserted.

- **Excel output template** (per sheet/tab):
  - Rows 1–4:
    - `A1 = Employer`
    - `A2 = Server`
    - `A3 = Test AccountID`
    - `A4 = Config #`
  - Row 5 (yellow header row):
    - `A5` (yellow background, empty text)
    - `B5 = Question Name:`
    - `C5 = Condition/Response`
    - `D5 = Expected Outcome:`
    - `E5 = Actual Outcome:`
    - `F5 = Comments:`
    - `G5 = Reviewed By:`
    - `H5 = Date`
    - `I5 = Comments`
  - Row 6 (orange, empty row).
  - From row 7 onward:
    - Column A: left blank.
    - Column B: question text (derived from the `Flag` column).
    - Column C: condition/response text.
    - Column D: **Expected Outcome** (pre-filled for some questions, empty for others).
    - Column E: **Actual Outcome** (initially empty for QA to fill).
    - Columns F–I: left empty for manual use (comments, reviewer, date, etc.).
  - Column widths:
    - A: 20, B: 70, C: 30, D–I: 20.

- **Multiple projects → multiple tabs**:
  - Every time the tool encounters `"This is the compliance matrix that has been applied to"`
    in a slide, it starts a new section and creates a new Excel sheet for that section.

- **.ppt support (optional)**:
  - If the uploaded file is **`.ppt`** (legacy binary format) and **LibreOffice** is available
    on the server (via `soffice` or `libreoffice` on PATH), the backend will convert it to
    `.pptx` first, then process it.
  - If LibreOffice is not installed, users should save the `.ppt` as `.pptx` in PowerPoint
    and upload the `.pptx` instead.

- **Dropdowns and statuses**:
  - A hidden sheet `Lookups` defines a named range `Outcome` pointing at `Lookups!$C$3:$C$8`,
    containing:
    - (blank)
    - `Green`
    - `Yellow`
    - `Red`
    - `No Flag`
    - `N/A`
  - For each data row starting from row 7 where **column C is not empty**, columns **D** and **E**
    get a **list data validation**: `type=list`, `formula1=Outcome`, `allowBlank=True`,
    `showDropDown=False` (Excel shows a dropdown arrow).
  - Both D and E have conditional formatting so that when the user selects `Green/Yellow/Red`,
    the font and background colors match the original template’s style.

### Requirements

1. **Python 3.10+** installed locally.
2. (Optional but recommended) a virtual environment.
3. For `.ppt` support: **LibreOffice** installed and the `soffice` (or `libreoffice`) binary
   available on PATH.

### Local setup

In your terminal:

```bash
cd /Users/andyzhu/WorkProject
python -m venv .venv
source .venv/bin/activate  # On Windows use: .venv\Scripts\activate
pip install -r requirements.txt
```

### Running the backend

From the project root (with the virtual environment activated):

```bash
uvicorn backend.main:app --reload
```

By default this starts the app at `http://127.0.0.1:8000/`:

- `GET /` – serves the static frontend (upload page for PPTX/PPT).
- `POST /api/convert` – accepts a PPT/PPTX file and returns the generated Excel file.

### How to use (local test)

1. Start the server:

   ```bash
   uvicorn backend.main:app --reload
   ```

2. Open the browser and navigate to:

   ```text
   http://127.0.0.1:8000/
   ```

3. On the page:
   - Click **“Choose PPTX file…”** and select a `.pptx` file that follows your compliance matrix template.
   - Click **“Convert to test sheet”**.
   - Wait a few seconds; the browser will automatically download the generated Excel file.

4. Open the Excel and verify:
   - Rows 1–6 match the header template.
   - From row 7 onward, questions and conditions are populated as expected.
   - Grey separator rows from PPT appear as grey empty rows in Excel.
   - Gate rows for `Please note ... YES/NO` appear as the fixed texts
     `If anwered Yes to above queestion` / `If anwered No to above queession`.
   - Columns D and E have the dropdown list (blank / Green / Yellow / Red / No Flag / N/A)
     wherever column C has a value.

### Adjusting behavior for other PPT templates

Depending on your real PPT files, you may want to tweak:

- **Grey row detection**:
  - Currently: a row is grey if all cells are text-empty and at least one cell has a
    grey-ish solid fill (RGB channels close to each other and intermediate brightness).
  - If your PPT uses different greys, adjust `_is_grey_empty_row` in `backend/converter.py`.

- **`Flag` column detection**:
  - Currently: any row with a cell text equal to `Flag` (case-insensitive) is treated as header,
    and that column index is used as the question text source.
  - If your template uses a different label, change `_find_flag_header_row`.

- **Condition detection logic**:
  - The mapping for Yes/No style questions and for general `Approved / Approved with Restriction / Not Approved`
    rows lives in `_append_rows_for_table_row` in `backend/converter.py`.
  - You can change the heuristics or add new patterns if your PPT uses different wording.

If you find real PPT files that are not converted as expected, you can adjust the rules based on
those examples—for example, changing how questions are grouped, adding more gate texts, or
outputting extra columns such as module names or test IDs.


