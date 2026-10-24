import argparse
import csv
import math
import re
import subprocess
import sys
from pathlib import Path

import numpy as np
import xlwings as xw


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = next(
    (parent for parent in SCRIPT_DIR.parents if (parent / "pipeline_logging.py").is_file()),
    SCRIPT_DIR.parent,
)
sys.path.insert(0, str(PROJECT_DIR))

from pipeline_logging import find_iitpave_directory, get_logger  # noqa: E402


LOGGER = get_logger(__file__)
IITPAVE_DIRECTORY = find_iitpave_directory(PROJECT_DIR)
IITPAVE_EXECUTABLE = IITPAVE_DIRECTORY / "IITPFILE.exe"
IITPAVE_INPUT = IITPAVE_DIRECTORY / "IITPAVE.IN"
IITPAVE_OUTPUT = IITPAVE_DIRECTORY / "IITPAVE.out"

INPUT_FIELDS = {
    "chainage": (
        "chainage",
        "chainage km",
        "chainage (km)",
        "a",
    ),
    "bituminous_thickness": (
        "bituminous layer thickness",
        "bituminous layer thickness mm",
        "bituminous layer thickness (mm)",
        "m",
    ),
    "base_thickness": (
        "base layer thickness",
        "base layer thickness mm",
        "base layer thickness (mm)",
        "n",
    ),
    "bituminous_modulus": (
        "bituminous e1",
        "bituminous (e1)",
        "bituminous modulus",
        "u",
    ),
    "base_modulus": (
        "base e2",
        "base (e2)",
        "base modulus",
        "v",
    ),
    "subgrade_modulus": (
        "sub grade e3",
        "sub grade (e3)",
        "subgrade e3",
        "subgrade modulus",
        "w",
    ),
    "asphalt_temperature": (
        "asphalt temperature",
        "asphalt temperature c",
        "asphalt temperature (c)",
        "ah",
    ),
    "season": (
        "season",
        "aj",
    ),
}

OUTPUT_HEADERS = [
    "Chainage (km)",
    "Bituminous Layer Thickness (mm)",
    "Base Layer Thickness (mm)",
    "Bituminous (E1)",
    "Base (E2)",
    "Sub-Grade (E3)",
    "Asphalt Temperature",
    "Season",
    "Lambda",
    "Corrected Bituminous E",
    "Corrected Base E",
    "Corrected Sub-Grade E",
    "Section No.",
    "From (km)",
    "To (km)",
    "Length (km)",
    "15th Percentile Bituminous E",
    "15th Percentile Base E",
    "15th Percentile Sub-Grade E",
    "Fatigue Strain",
    "Rutting Strain",
    "Fatigue Residual Life (MSA)",
    "Rutting Residual Life (MSA)",
    "Remaining Life (MSA)",
]


def _normalize_header(value):
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()


def _header_columns(header_row):
    normalized = {}
    for column_index, value in enumerate(header_row):
        key = _normalize_header(value)
        if key:
            normalized.setdefault(key, column_index)

    columns = {}
    missing = []
    for field, aliases in INPUT_FIELDS.items():
        column_index = next(
            (
                normalized[_normalize_header(alias)]
                for alias in aliases
                if _normalize_header(alias) in normalized
            ),
            None,
        )
        if column_index is None:
            missing.append(field)
        else:
            columns[field] = column_index
    if missing:
        raise ValueError(
            "Input is missing required column(s): "
            + ", ".join(missing)
            + ". See the RMN_LF.py instructions for accepted column names."
        )
    return columns


def _read_excel_rows(input_path):
    with xw.App(visible=False) as app:
        workbook = app.books.open(str(input_path), read_only=True, update_links=False)
        try:
            sheet = workbook.sheets[0]
            used_range = sheet.used_range
            values = used_range.value
            if not isinstance(values, list):
                raise ValueError("The Excel input sheet is empty.")
            if values and not isinstance(values[0], list):
                values = [values]

            for header_index, row in enumerate(values[:10]):
                try:
                    columns = _header_columns(row)
                except ValueError:
                    continue
                return values[header_index + 1 :], columns, header_index + 2
        finally:
            workbook.close()
    raise ValueError(
        "Could not find a header row containing all required fields in the first "
        "10 rows of the Excel input."
    )


def _read_text_rows(input_path):
    text = None
    for encoding in ("utf-8-sig", "utf-16"):
        try:
            text = input_path.read_text(encoding=encoding)
            break
        except UnicodeError:
            continue
    if text is None:
        raise UnicodeError(f"Could not decode text input file: {input_path}")

    sample = text[:8192]
    try:
        delimiter = csv.Sniffer().sniff(sample, delimiters="\t,;").delimiter
    except csv.Error:
        delimiter = "\t" if "\t" in sample else ","
    rows = list(csv.reader(text.splitlines(), delimiter=delimiter))
    for header_index, row in enumerate(rows[:10]):
        try:
            columns = _header_columns(row)
        except ValueError:
            continue
        return rows[header_index + 1 :], columns, header_index + 2
    raise ValueError(
        "Could not find a delimited-text header row containing all required fields "
        "in the first 10 rows."
    )


def load_input_rows(input_path):
    suffix = input_path.suffix.lower()
    if suffix in (".xlsx", ".xlsm", ".xls"):
        rows, columns, first_data_line = _read_excel_rows(input_path)
    elif suffix in (".txt", ".csv", ".tsv"):
        rows, columns, first_data_line = _read_text_rows(input_path)
    else:
        raise ValueError(
            f"Unsupported input type {input_path.suffix!r}; use Excel (.xlsx/.xlsm) "
            "or delimited text (.txt/.csv/.tsv)."
        )

    records = []
    for line_number, row in enumerate(rows, start=first_data_line):
        if not row or all(value is None or str(value).strip() == "" for value in row):
            continue
        record = {}
        for field, column_index in columns.items():
            raw_value = row[column_index] if column_index < len(row) else None
            if field == "season":
                record[field] = str(raw_value).strip() if raw_value is not None else ""
                continue
            try:
                number = float(raw_value)
            except (TypeError, ValueError):
                raise ValueError(
                    f"Invalid {field.replace('_', ' ')} at input row {line_number}: "
                    f"{raw_value!r}"
                )
            if not math.isfinite(number):
                raise ValueError(
                    f"Non-finite {field.replace('_', ' ')} at input row "
                    f"{line_number}: {raw_value!r}"
                )
            record[field] = number

        if record["chainage"] < 0:
            raise ValueError(f"Chainage at input row {line_number} cannot be negative.")
        if record["asphalt_temperature"] <= 0:
            raise ValueError(
                f"Asphalt temperature at input row {line_number} must be positive."
            )
        if any(
            record[field] <= 0
            for field in (
                "bituminous_thickness",
                "base_thickness",
                "bituminous_modulus",
                "base_modulus",
                "subgrade_modulus",
            )
        ):
            raise ValueError(
                f"Layer thicknesses and predicted moduli must be positive at "
                f"input row {line_number}."
            )
        if not record["season"]:
            raise ValueError(f"Season is blank at input row {line_number}.")
        records.append(record)

    if not records:
        raise ValueError("No input data rows were found.")
    return records


def _read_boundaries():
    while True:
        boundary_text = input(
            "Enter variable chainage boundaries in km, comma-separated "
            "(start with 0 and finish with end; e.g. 0, 1, 7.34, 15, end): "
        ).strip()
        try:
            tokens = [token.strip().lower() for token in boundary_text.split(",")]
            if len(tokens) < 2 or tokens[0] != "0" or tokens[-1] != "end":
                raise ValueError(
                    "Boundaries must start with 0 and the final entry must be 'end'."
                )
            boundaries = [float(token) for token in tokens[:-1]]
            if any(not math.isfinite(value) for value in boundaries):
                raise ValueError("All numeric boundaries must be finite numbers.")
            if any(
                upper <= lower for lower, upper in zip(boundaries, boundaries[1:])
            ):
                raise ValueError(
                    "Boundaries must be strictly increasing with no duplicates."
                )
            return boundaries
        except ValueError as exc:
            print(f"Invalid boundaries: {exc}")


def _correct_row(record):
    temperature = record["asphalt_temperature"]
    denominator = 1 - 0.238 * math.log(temperature)
    if math.isclose(denominator, 0.0, abs_tol=1e-12):
        raise ValueError(
            "Asphalt temperature makes the temperature correction undefined: "
            f"{temperature}."
        )
    correction_factor = (1 - 0.238 * math.log(35)) / denominator
    base_modulus = record["base_modulus"]
    subgrade_modulus = record["subgrade_modulus"]
    season = record["season"].strip().lower()

    if season == "summer":
        corrected_base = -0.0003 * (base_modulus**2) + 0.9854 * base_modulus - 32.989
        corrected_subgrade = 0.8554 * subgrade_modulus - 8.461
    elif season == "winter":
        corrected_base = 10.5523 * (base_modulus**0.624) - 113.857
        corrected_subgrade = 3.351 * (subgrade_modulus**0.7688) - 28.9
    else:
        corrected_base = base_modulus
        corrected_subgrade = subgrade_modulus

    record["lambda"] = correction_factor
    record["corrected_bituminous"] = correction_factor * record["bituminous_modulus"]
    record["corrected_base"] = corrected_base
    record["corrected_subgrade"] = corrected_subgrade


def _run_iitpave(moduli, thicknesses):
    input_content = (
        f"3\n"
        f"{' '.join(map(str, moduli))}\n"
        f"0.5 0.4 0.4\n"
        f"{' '.join(map(str, thicknesses))}\n"
        "20000 0.56\n"
        "4\n"
        f"{thicknesses[0]} 0\n"
        f"{thicknesses[0]} 152\n"
        f"{thicknesses[0] + thicknesses[1]} 0\n"
        f"{thicknesses[0] + thicknesses[1]} 152\n"
        "2\n"
    )
    IITPAVE_INPUT.write_text(input_content, encoding="utf-8")
    LOGGER.info("Running IITPAVE: %s", IITPAVE_EXECUTABLE)
    subprocess.run(
        [str(IITPAVE_EXECUTABLE)],
        cwd=IITPAVE_DIRECTORY,
        check=True,
        capture_output=True,
        text=True,
    )
    if not IITPAVE_OUTPUT.is_file():
        raise FileNotFoundError(f"IITPAVE output was not created: {IITPAVE_OUTPUT}")
    return IITPAVE_OUTPUT


def _get_fatigue_strain(file_path):
    number_pattern = r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[Ee][+-]?\d+)?"
    values = []
    parsing = False
    with file_path.open("r", encoding="utf-8", errors="replace") as output_file:
        for line in output_file:
            if "epT" in line:
                parsing = True
                continue
            if parsing:
                numbers = re.findall(number_pattern, line)
                if len(numbers) >= 9:
                    values.append(abs(float(numbers[8])))
    if len(values) < 3:
        raise ValueError(
            f"Expected at least 3 epT strain readings in {file_path}; found {len(values)}"
        )
    return max(values[0], values[2])


def _get_rutting_strain(file_path):
    number_pattern = r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[Ee][+-]?\d+)?"
    values = []
    parsing = False
    with file_path.open("r", encoding="utf-8", errors="replace") as output_file:
        for line in output_file:
            if "epZ" in line:
                parsing = True
                continue
            if parsing:
                numbers = re.findall(number_pattern, line)
                if len(numbers) >= 8:
                    values.append(abs(float(numbers[7])))
    if len(values) < 8:
        raise ValueError(
            f"Expected at least 8 epZ strain readings in {file_path}; found {len(values)}"
        )
    return max(values[5], values[7])


def calculate_remaining_life(records, numeric_boundaries):
    for record in records:
        _correct_row(record)

    max_chainage = max(record["chainage"] for record in records)
    if max_chainage <= numeric_boundaries[-1]:
        raise ValueError(
            f"The last numeric boundary ({numeric_boundaries[-1]:.6g} km) must be "
            f"less than the maximum chainage ({max_chainage:.6g} km)."
        )

    boundaries = numeric_boundaries + [max_chainage]
    section_number = 0
    for lower_km, upper_km in zip(boundaries, boundaries[1:]):
        section_records = [
            record
            for record in records
            if (
                lower_km <= record["chainage"] <= upper_km
                if lower_km == 0
                else lower_km < record["chainage"] <= upper_km
            )
        ]
        if not section_records:
            LOGGER.warning(
                "No readings found in requested interval %.6g-%.6g km; skipping.",
                lower_km,
                upper_km,
            )
            continue

        section_number += 1
        bituminous_values = [row["corrected_bituminous"] for row in section_records]
        base_values = [row["corrected_base"] for row in section_records]
        subgrade_values = [row["corrected_subgrade"] for row in section_records]
        bituminous_e, base_e, subgrade_e = (
            float(np.quantile(values, 0.15))
            for values in (bituminous_values, base_values, subgrade_values)
        )
        first_record = section_records[0]
        thicknesses = [
            first_record["bituminous_thickness"],
            first_record["base_thickness"],
        ]
        if any(
            (row["bituminous_thickness"], row["base_thickness"])
            != tuple(thicknesses)
            for row in section_records
        ):
            LOGGER.warning(
                "Section %d contains varying layer thicknesses; IITPAVE will use "
                "the first input row's thicknesses %r.",
                section_number,
                thicknesses,
            )

        output_file = _run_iitpave([bituminous_e, base_e, subgrade_e], thicknesses)
        fatigue_strain = _get_fatigue_strain(output_file)
        rutting_strain = _get_rutting_strain(output_file)
        if bituminous_e <= 0 or fatigue_strain <= 0 or rutting_strain <= 0:
            raise ValueError(
                f"Section {section_number}: modulus and strains must be positive; "
                f"bituminous E={bituminous_e}, fatigue strain={fatigue_strain}, "
                f"rutting strain={rutting_strain}."
            )

        fatigue_life = (
            (0.711 * 10**-4)
            * ((1 / fatigue_strain) ** 3.89)
            * ((1 / bituminous_e) ** 0.854)
            * (10**-6)
        )
        rutting_life = (1.41 * 10**-8) * ((1 / rutting_strain) ** 4.5337) * (10**-6)
        first_record["section_results"] = [
            section_number,
            lower_km,
            upper_km,
            upper_km - lower_km,
            bituminous_e,
            base_e,
            subgrade_e,
            fatigue_strain,
            rutting_strain,
            fatigue_life,
            rutting_life,
            min(fatigue_life, rutting_life),
        ]
        LOGGER.info(
            "Section %d completed: %.6g-%.6g km; remaining life=%g MSA",
            section_number,
            lower_km,
            upper_km,
            first_record["section_results"][-1],
        )


def write_output(records, output_path):
    rows = []
    for record in records:
        row = [
            record["chainage"],
            record["bituminous_thickness"],
            record["base_thickness"],
            record["bituminous_modulus"],
            record["base_modulus"],
            record["subgrade_modulus"],
            record["asphalt_temperature"],
            record["season"],
            record["lambda"],
            record["corrected_bituminous"],
            record["corrected_base"],
            record["corrected_subgrade"],
        ]
        row.extend(record.get("section_results", [None] * 12))
        rows.append(row)

    with xw.App(visible=False) as app:
        app.display_alerts = False
        workbook = app.books.add()
        try:
            sheet = workbook.sheets[0]
            sheet.name = "Remaining Life"
            sheet.range("A1").value = [OUTPUT_HEADERS]
            sheet.range(f"A2:X{len(rows) + 1}").value = rows
            header = sheet.range("A1:X1")
            header.api.Font.Bold = True
            header.api.WrapText = True
            sheet.range("A:X").api.Columns.AutoFit()
            sheet.api.Rows(1).RowHeight = 36
            workbook.save(str(output_path))
        finally:
            workbook.close()


def main():
    parser = argparse.ArgumentParser(
        description="Calculate pavement remaining life from a minimal Excel or text input."
    )
    parser.add_argument(
        "input_file",
        nargs="?",
        help="Input .xlsx/.xlsm/.xls or delimited .txt/.csv/.tsv file",
    )
    args = parser.parse_args()
    provided_name = args.input_file or input(
        "Enter the input filename (the file must be in the Remaining Life folder): "
    ).strip().strip('"')
    input_name = Path(provided_name).name
    if input_name != provided_name:
        raise ValueError("Enter only the input filename, not a folder path.")
    input_path = (SCRIPT_DIR / input_name).resolve()
    if not input_path.is_file():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    records = load_input_rows(input_path)
    LOGGER.info("Loaded %d data rows from %s", len(records), input_path)
    boundaries = _read_boundaries()
    calculate_remaining_life(records, boundaries)

    output_path = input_path.with_name(f"{input_path.stem}_RMN-LF.xlsx")
    write_output(records, output_path)
    LOGGER.info("Remaining-life output saved to %s", output_path)
    print(f"Remaining-life output saved to: {output_path}")


if __name__ == "__main__":
    main()
