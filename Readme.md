# Pavement Analysis Automation Scripts

This folder contains a set of Python scripts used for pavement layer analysis, modulus back-calculation, remaining-life prediction, overlay requirement checks, and RMSE validation. The scripts work with Microsoft Excel workbooks and external pavement analysis tools (IITPAVE and KGPBACK) to automate engineering calculations for road sections.

The overall purpose is to:
- read pavement deflection data from .fwd files,
- calculate or back-calculate layer moduli,
- compare measured and predicted deflections,
- determine whether overlays are required,
- estimate remaining life and fatigue/rutting responses,
- write results back into Excel sheets for reporting and design.

These scripts are Windows-oriented and rely on Excel automation via `xlwings`, as well as external executables and analysis output files.

---

## 1) INPUT.py

### What it does
This script reads a pavement deflection file (`MANOHARPUR TO DAUSA RHS.fwd`) and prepares an Excel workbook named `Modulus_Calculation.xlsx` for further analysis. It accepts both whitespace/tab-separated measurement rows and rows prefixed with `D`; it detects UTF-8 and UTF-16 input.

### Main tasks
- Opens and parses the FWD data file, skipping headers, blank lines, and malformed rows.
- Extracts deflection values and related metadata such as:
  - chainage,
  - impact number,
  - load,
  - sensor deflections D0 to D8 when present (legacy seven-sensor rows leave D7/D8 as `NA`),
  - air temperature,
  - pavement temperature,
  - modulus label,
  - time.
- Groups measurements at the same chainage and averages load-normalized sensor deflections.
- Creates or opens the workbook and adds a sheet named `RHS`.
- Writes a large formatted Excel table with column headers and section titles.
- Fills rows with chainage, load, stress, sensor deflections, thicknesses, E-value ranges, predicted modulus values, RMSE, temperature, lambda, season, and residual-life-related columns.
- Uses `xlwings` formatting commands to merge cells, apply bold text, add borders, and style the sheet.
- Logs the input file, number of valid/skipped lines, output rows, and workbook save location to both the console and the shared `pipeline.log`.

### Engineering logic
It converts measured deflections into a normalized form and stores them in the Excel workbook so other scripts can use them for back-calculation and overlay analysis.

### Important note
This script does not do pavement modeling itself; it prepares the workbook and deflection data for later analysis. The input file is read only during a normal import. The current `MANOHARPUR TO DAUSA RHS.fwd` has explicitly marked synthetic D7/D8 example values added for importer testing; those two sensors are not represented as verified field measurements.

### Running and checking the import

From this folder, run:

```powershell
python INPUT.py
```

The defaults read `MANOHARPUR TO DAUSA RHS.fwd` and write to `Modulus_Calculation.xlsx`. To keep the existing workbook untouched, pass a different output workbook path:

```powershell
python INPUT.py --input "MANOHARPUR TO DAUSA RHS.fwd" --workbook Modulus_Calculation_CHECK.xlsx
```

Open the selected workbook in Excel and inspect the `RHS` worksheet. Imported points begin on row 3; chainage is in column A and normalized sensor deflections are in columns D:L when all nine sensor values are present. Older seven-sensor rows leave the last two sensor cells as `NA`. The run log reports how many input records were accepted and which chainages were written. Close the workbook in Excel or WPS before running the importer; it now reports an error rather than claiming success if the workbook is opened read-only.

The current `MANOHARPUR TO DAUSA RHS.fwd` includes synthetic D7/D8 values specifically for testing the import layout and decreasing sensor readings. They are labelled in the input file and must not be treated as measured field data.

## Shared run log

All Python scripts write timestamped INFO messages and uncaught exception tracebacks to `pipeline.log` in this folder. The log also records when a script skips workbook rows because required input cells are blank. Each script still writes progress to the terminal. Close `Modulus_Calculation.xlsx` in Excel/WPS before running a script that updates it; scripts report an error instead of silently proceeding if the workbook is read-only.

---

## 2) KGPBack.py

### What it does
This script uses the project-folder `KGPBACK.exe` to back-calculate pavement layer moduli from deflection basin data stored in `Modulus_Calculation.xlsx`.

### Main tasks
- Opens `Modulus_Calculation.xlsx`.
- Reads rows in the `RHS` worksheet.
- For each pavement point, it extracts:
  - measured deflections from D columns,
  - layer thicknesses from M and N columns,
  - E-value ranges from O to T columns.
- Builds the input file text expected by the `KGPBACK.exe` program.
- Runs the executable via `subprocess.Popen`.
- Reads the output file `backout` and extracts modulus values for:
  - Surface (MPa),
  - Base (MPa),
  - Subgrade (MPa).
- Writes these values into columns `U`, `V`, and `W` in the workbook.

### Helper functions
- `run_kgpback_jar(defln, thick, E_range, pos)`: creates the input content and runs the external program.
- `get_dispz_regex(file_path)`: scans the output text and extracts the three modulus values using a regex-based numeric parser.

### Purpose in the workflow
This is the back-calculation step: it takes measured deflections and uses a separate program to estimate pavement layer stiffness values.

---

## 3) OVERLAY_REQ.py

### What it does
This script determines whether overlay is required and calculates the overlay thickness needed to meet a minimum design life requirement. It reads the `RHS` worksheet in `Modulus_Calculation.xlsx` and requires remaining-life values in column AP.

### Main tasks
- Opens `Modulus_Calculation.xlsx` from this folder.
- Iterates through pavement sections or rows.
- Uses IITPAVE (via `IITPFILE.exe`) to compute fatigue and rutting strains for various overlay thicknesses.
- Evaluates whether the remaining life is below a required MSA threshold.
- If the life is insufficient, it sets `AQ` to `Yes` and finds the minimum overlay thickness required.
- Writes results into columns such as:
  - `AQ` = Overlay Required (Yes/No)
  - `AR` = Overlay thickness required (mm)
  - `AS` = Overlay modulus
  - `AT` = Poisson's ratio of overlay
  - `AU`, `AV` = fatigue and rutting strain after overlay
  - `AW`, `AX` = residual life values
  - `AY` = final remaining life after overlay

### Calculation logic
- For rows with insufficient remaining life, it loops from 5 mm to 1000 mm in 5 mm increments.
- It creates a pavement model using the current moduli and layer thicknesses.
- It runs IITPAVE with different overlay thicknesses.
- It reads fatigue and rutting strains from the output file.
- It computes life equations using formulas involving strain and modulus values.
- The script stops at the first thickness that meets the service-life requirement.

### Important features
- Uses a threshold check based on chainage and MSA (million standard axles).
- Handles both lower- and higher-chainage sections separately.
- Contains engineering formulas for fatigue and rutting life.

---

## 4) REMAINING_LIFE.py

### What it does
This script estimates the remaining life of pavement layers after applying seasonal and temperature corrections, then groups sections into homogeneous lengths for analysis. It reads and writes the populated `RHS` worksheet in `Modulus_Calculation.xlsx`.

### Main tasks
- Opens the workbook `Modulus_Calculation (Manoharpur to Dausa).xlsx`.
- Reads the `RHS` worksheet (`ws3`).
- Applies correction factors to modulus values using seasonal and temperature equations.
- Writes corrected values to columns `Z`, `AB`, `AC`, and `AD`.
- Groups rows into homogeneous pavement sections based on matching thicknesses and chainage continuity.
- For each homogeneous segment it calculates:
  - 15th percentile modulus values for each layer,
  - fatigue strain,
  - rutting strain,
  - fatigue life, rutting life,
  - minimum remaining life.
- Writes these values to columns `AE` through `AP`.

### Key helper functions
- `run_iitpave_jar(moduli, thicknesses, number, poi)`: calls the external `IITPFILE.exe` executable with the required input.
- `get_fatigue_strain(file_path)`: extracts the maximum fatigue strain from the IITPAVE output.
- `get_rutting_strain(file_path)`: extracts the maximum rutting strain from the output.

### Engineering logic
- It does not simply process each row independently. It merges neighboring rows into homogeneous sections with similar layer thicknesses.
- For each section it computes the 15th percentile modulus value to represent a conservative lower bound.
- Then it runs the pavement model and calculates residual life with formulas tied to fatigue and rutting behavior.
- This is a structural design and pavement performance evaluation step.

---

## 5) Remaining Life/RMN_LF.py

This standalone version calculates corrected moduli and remaining-life results from a compact input file, without requiring the full `Modulus_Calculation.xlsx` worksheet. It runs IITPAVE independently for every input row and writes that row's results to the corresponding output row. The output is saved beside the selected source file as `<input basename>_RMN-LF.xlsx` (for example, `survey_RMN-LF.xlsx` from `survey.xlsx`).

Run it from the repository root or from the `Remaining Life` folder:

```powershell
python "Remaining Life\RMN_LF.py" "Remaining Life\your-input.xlsx"
```

If no path is supplied, the script prompts for one. Excel inputs may be `.xlsx`, `.xlsm`, or `.xls`; text inputs may be tab-, comma-, or semicolon-delimited `.txt`, `.csv`, or `.tsv` files. Both formats need a header row and these eight fields:

| Field | Accepted examples |
| --- | --- |
| Chainage in kilometres | `Chainage`, `Chainage (km)`, `A` |
| Bituminous layer thickness in mm | `Bituminous Layer Thickness (mm)`, `M` |
| Base layer thickness in mm | `Base Layer Thickness (mm)`, `N` |
| Predicted bituminous modulus | `Bituminous (E1)`, `U` |
| Predicted base modulus | `Base (E2)`, `V` |
| Predicted sub-grade modulus | `Sub-Grade (E3)`, `W` |
| Asphalt temperature | `Asphalt Temperature`, `AH` |
| Season | `Season`, `AJ` |

After reading the file, the script asks for chainage boundaries in kilometres, as the original remaining-life workflow does. IITPAVE and Microsoft Excel are required to perform the calculation and write the `.xlsx` result.

## 6) RMSE.py

### What it does
This script calculates the Root Mean Square Error (RMSE) between measured deflection data and theoretical deflections predicted by IITPAVE for a given set of moduli.

### Main tasks
- Opens `Modulus_Calculation.xlsx`.
- Reads the measured deflection basin values from columns D to L for each row.
- Extracts sensor positions and thickness values.
- Builds a moduli vector from the computed layer moduli.
- Calls `run_iitpave_jar(...)` to generate predicted deflection values.
- Reads the `DispZ` output from the IITPAVE output file.
- Computes the RMSE and percentage RMSE using:
  - `RMSE = sqrt(mean((theoretical - measured)^2))`
  - `PCT_RMSE = (RMSE / mean(measured)) * 100`
- Writes the percentage RMSE to column `X`.

### Helper functions
- `run_iitpave_jar(moduli, thicknesses, number, poi, pos)`: generates the input deck for IITPAVE and runs the program.
- `get_dispz_regex(file_path)`: extracts deflection values from the output file based on the `DispZ` keyword.
- `RMSE_function(moduli)`: the central function that computes the RMSE for a single section.

### Purpose
This script is used as a calibration or validation step: it checks how closely the selected moduli reproduce the measured deflection basin.

---

## Shared technical pattern across the scripts

All of these scripts follow a similar pattern:

1. Read data from an Excel workbook (`xlwings`).
2. Build a model input for an external pavement analysis engine.
3. Run an EXE or JAR file via `subprocess`.
4. Parse output files using regular expressions.
5. Write numerical results back into Excel columns.
6. Repeat row-by-row or section-by-section.

This is a classic engineering automation workflow: spreadsheet-driven pavement analysis with external computational tools.

---

## External tools and dependencies
These Python scripts depend on:
- `xlwings` for Excel automation
- `numpy` for array operations and percentile calculations
- `scipy.optimize.minimize` in some cases
- `re` for parsing output text
- `subprocess` for invoking external programs
- external executables like:
  - `IITPFILE.exe`
  - `KGPBACK.exe`
  - associated input/output files such as `IITPAVE.IN`, `IITPAVE.out`, `backout`

  The IITPAVE scripts locate `IITPFILE.exe` inside the extracted `IRC 37 2018-20261017T090045Z-1-001\IRC 37 2018\IITPAVE` project folder. Close `Modulus_Calculation.xlsx` in Excel/WPS before running scripts that save results.

  All five Python scripts write timestamped progress, warnings, errors, and uncaught exception tracebacks to the shared `pipeline.log` file in this folder.

---

## Overall workflow of the project

The project appears to be a pavement evaluation pipeline:

1. `INPUT.py` prepares the raw FWD deflection dataset into an Excel workbook.
2. `KGPBack.py` back-calculates layer moduli.
3. `RMSE.py` checks how well the chosen moduli fit measured deflections.
4. `REMANING_LIFE.py` computes residual life and homogenized segment behavior.
5. `OVERLAY_REQ.py` decides if an overlay is needed and estimates overlay thickness.

Together, these scripts form a full decision-support workflow for pavement structural assessment and rehabilitation design.

---

## Practical note
The scripts expect this project folder layout, the named workbook, and the external analysis executables. If the files are moved, confirm the workbook and executable paths before running.
