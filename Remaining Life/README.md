# Remaining Life Calculator (`RMN_LF.py`)

`RMN_LF.py` calculates temperature- and season-corrected moduli, evaluates pavement sections with IITPAVE, and reports fatigue life, rutting life, and minimum remaining life. It is intended for a compact input file containing only the fields needed for this calculation; it does not require the full `Modulus_Calculation.xlsx` workbook.

## Run the script

The input file must be in the same `Remaining Life` folder as this script. From the repository root, start it without an argument to be prompted for the **filename only**:

```powershell
python "Remaining Life\RMN_LF.py"
```

At the prompt, enter a name such as `survey.xlsx` (do not enter a folder path). The file will be resolved from the `Remaining Life` folder.

You can also supply just the filename as an optional argument:

```powershell
python "Remaining Life\RMN_LF.py" survey.xlsx
```

The output is saved in the `Remaining Life\Output` folder, which is created automatically if it does not already exist. The input filename's base name is used with `_RMN-LF` appended. For example:

```text
Remaining Life\survey.xlsx
    -> Remaining Life\Output\survey_RMN-LF.xlsx
```

The input file is read only; the script creates a separate output workbook.

## Input file formats

Supported input extensions:

- Excel: `.xlsx`, `.xlsm`, `.xls`
- Delimited text: `.txt`, `.csv`, `.tsv`

Excel input uses the **first worksheet**. Text input may be tab-, comma-, or semicolon-delimited. UTF-8 (with or without a byte-order mark) and UTF-16 text are supported.

The file must contain a header row and the eight required fields below. In Excel, the script reads the first worksheet and searches its first 10 used rows for a complete header row; in text input, it searches the first 10 lines. It then finds each field by its header name, regardless of its position in the input. Data values are read from the corresponding source column on each row immediately below the header. For example, if `Chainage (km)` is in source cell `D4` and the header is in row 3, the first chainage is read from `D4`. Input fields do not need to be in any particular worksheet column order. Blank rows are ignored.

The letter aliases in the table refer to the matching columns in the full `Modulus_Calculation.xlsx` workflow. When using those aliases in this compact input, put the letter itself in the header cell (for example, `A` or `AH`); the input column can still be in any position. These letters are header aliases, not required physical source-cell addresses.

| Required input field | Original worksheet column | Accepted header names (case-insensitive) | Units / expected values |
| --- | --- | --- | --- |
| Chainage | A | `Chainage`, `Chainage km`, `Chainage (km)`, `A` | Kilometres; numeric, zero or greater |
| Bituminous layer thickness | M | `Bituminous Layer Thickness`, `Bituminous Layer Thickness mm`, `Bituminous Layer Thickness (mm)`, `M` | Millimetres; positive number |
| Base layer thickness | N | `Base Layer Thickness`, `Base Layer Thickness mm`, `Base Layer Thickness (mm)`, `N` | Millimetres; positive number |
| Predicted bituminous modulus (E1) | U | `Bituminous E1`, `Bituminous (E1)`, `Bituminous Modulus`, `U` | Positive modulus value (MPa) |
| Predicted base modulus (E2) | V | `Base E2`, `Base (E2)`, `Base Modulus`, `V` | Positive modulus value (MPa) |
| Predicted sub-grade modulus (E3) | W | `Sub Grade E3`, `Sub Grade (E3)`, `Subgrade E3`, `Subgrade Modulus`, `W` | Positive modulus value (MPa) |
| Asphalt temperature | AH | `Asphalt Temperature`, `Asphalt Temperature C`, `Asphalt Temperature (C)`, `AH` | Degrees Celsius; positive number |
| Season | AJ | `Season`, `AJ` | Non-empty text; `Summer` and `Winter` apply their respective seasonal corrections |

Header matching ignores letter case and punctuation. For example, `Asphalt Temperature (C)` and `asphalt-temperature c` are treated equivalently. Data fields other than season must be numeric and finite. Layer thicknesses and moduli must be positive; chainage must be non-negative. Unrecognized season values receive no base or sub-grade seasonal correction.

### Minimal input example

| Chainage (km) | Bituminous Layer Thickness (mm) | Base Layer Thickness (mm) | Bituminous (E1) | Base (E2) | Sub-Grade (E3) | Asphalt Temperature | Season |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| 0.5 | 115 | 450 | 3000 | 500 | 100 | 35 | Summer |
| 1.0 | 115 | 450 | 3100 | 520 | 110 | 35 | Summer |

These columns can appear in any order, provided their headers identify the required fields.

## Chainage boundaries

After loading and validating the input, the script prompts for chainage section boundaries in kilometres. Enter a comma-separated list that starts with `0` and ends with `end`, for example:

```text
0, 1, 7.34, end
```

The last `end` is resolved to the maximum chainage in the input. Numeric boundaries must be finite and strictly increasing, and the last numeric boundary must be below the maximum input chainage. Readings at a boundary belong to the interval ending at that boundary; the first interval includes chainage zero.

## Output workbook layout

The generated workbook contains a worksheet named **Remaining Life**. Headers are written to `A1:X1`; input data rows are written to `A2:X<n>`, where `<n>` is the last output row. Within each output data row, columns A:H contain copied input values and I:L contain row-level corrections. Section-level results are written to columns M:X only on the first output row belonging to each populated chainage interval; the section-result cells on other rows remain blank. For example, if a section's first input record is the third record, its section outputs go to `M4:X4` (output row 1 is the header).

| Output cell / column | Header | Value written |
| --- | --- | --- |
| A | Chainage (km) | Input chainage |
| B | Bituminous Layer Thickness (mm) | Input bituminous thickness |
| C | Base Layer Thickness (mm) | Input base thickness |
| D | Bituminous (E1) | Input predicted bituminous modulus |
| E | Base (E2) | Input predicted base modulus |
| F | Sub-Grade (E3) | Input predicted sub-grade modulus |
| G | Asphalt Temperature | Input asphalt temperature |
| H | Season | Input season |
| I | Lambda | Temperature correction factor: `(1 - 0.238 × ln(35)) / (1 - 0.238 × ln(temperature))` |
| J | Corrected Bituminous E | Input bituminous modulus × Lambda |
| K | Corrected Base E | Summer: `-0.0003 × E2² + 0.9854 × E2 - 32.989`; Winter: `10.5523 × E2^0.624 - 113.857`; otherwise, input E2 |
| L | Corrected Sub-Grade E | Summer: `0.8554 × E3 - 8.461`; Winter: `3.351 × E3^0.7688 - 28.9`; otherwise, input E3 |
| M | Section No. | Sequential number for the populated interval |
| N | From (km) | Lower boundary of the section |
| O | To (km) | Upper boundary of the section |
| P | Length (km) | Section upper boundary minus lower boundary |
| Q | 15th Percentile Bituminous E | 15th percentile of corrected bituminous modulus values in the section |
| R | 15th Percentile Base E | 15th percentile of corrected base modulus values in the section |
| S | 15th Percentile Sub-Grade E | 15th percentile of corrected sub-grade modulus values in the section |
| T | Fatigue Strain | Fatigue strain read from the IITPAVE output |
| U | Rutting Strain | Rutting strain read from the IITPAVE output |
| V | Fatigue Residual Life (MSA) | Fatigue-life calculation using fatigue strain and the 15th percentile bituminous modulus |
| W | Rutting Residual Life (MSA) | Rutting-life calculation using rutting strain |
| X | Remaining Life (MSA) | Minimum of fatigue life and rutting life |

The section-level values in columns M:X are located on the first data row in that interval, not repeated on every row in the section. IITPAVE uses the first row's two thickness values if thickness varies within a section.

## Calculation and dependencies

For each requested populated interval, the script uses the 15th percentile of each corrected modulus, runs `IITPFILE.exe`, extracts fatigue and rutting strains from `IITPAVE.out`, and computes:

- Fatigue residual life: `(0.711 × 10^-4) × (1 / fatigue strain)^3.89 × (1 / 15th-percentile bituminous E)^0.854 × 10^-6`
- Rutting residual life: `(1.41 × 10^-8) × (1 / rutting strain)^4.5337 × 10^-6`
- Remaining life: the lower of fatigue residual life and rutting residual life

Running the full calculation requires Windows, Microsoft Excel, the Python packages `xlwings` and `numpy`, and the IITPAVE `IITPFILE.exe` installation located by the project's shared IITPAVE directory lookup.
