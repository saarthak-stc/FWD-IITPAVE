import subprocess
import os
import xlwings as xw
import re
import math
import numpy as np

from pathlib import Path
from pipeline_logging import find_iitpave_directory, get_logger

SCRIPT_DIR = Path(__file__).resolve().parent
LOGGER = get_logger(__file__)


def run_iitpave_jar(moduli, thicknesses, number, poi):
    """Run IITPAVE for the selected layer moduli and thicknesses."""
    jar_path = os.path.join(JAR_DIRECTORY, JAR_NAME)
    input_path = os.path.join(JAR_DIRECTORY, INPUT_FILE_NAME)
    output_path = os.path.join(JAR_DIRECTORY, OUTPUT_FILE_NAME)

    if number == 3:
        input_content = f"""{" ".join(map(str, [number]))}
{" ".join(map(str, moduli))}
{" ".join(map(str, poi))}
{" ".join(map(str, thicknesses))}
20000 0.56
4
{thicknesses[0]} 0
{thicknesses[0]} 152
{thicknesses[0] + thicknesses[1]} 0
{thicknesses[0] + thicknesses[1]} 152
2
"""
    else:
        input_content = f"""{" ".join(map(str, [number]))}
{" ".join(map(str, moduli))}
{" ".join(map(str, poi))}
{" ".join(map(str, thicknesses))}
20000 0.56
4
{thicknesses[0] + thicknesses[1]} 0
{thicknesses[0] + thicknesses[1]} 152
{thicknesses[0] + thicknesses[1] + thicknesses[2]} 0
{thicknesses[0] + thicknesses[1] + thicknesses[2]} 152
2
"""

    with open(input_path, "w", encoding="utf-8") as file:
        file.write(input_content)

    try:
        LOGGER.info("Running IITPAVE: %s", jar_path)
        subprocess.run(
            [jar_path],
            cwd=JAR_DIRECTORY,
            check=True,
            capture_output=True,
            text=True,
        )
    except (subprocess.CalledProcessError, FileNotFoundError):
        LOGGER.exception("IITPAVE execution failed")
        raise

    if not os.path.isfile(output_path):
        raise FileNotFoundError(f"IITPAVE output was not created: {output_path}")

    return output_path


def get_fatigue_strain(file_path):
    num_pattern = r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[Ee][+-]?\d+)?"
    values = []
    with open(file_path, "r", encoding="utf-8", errors="replace") as file:
        start_parsing = False
        for line in file:
            if "epT" in line:
                start_parsing = True
                continue
            if start_parsing:
                numbers = re.findall(num_pattern, line)
                if len(numbers) >= 9:
                    values.append(abs(float(numbers[8])))

    if len(values) < 3:
        raise ValueError(
            f"Expected at least 3 epT strain readings in {file_path}; found {len(values)}"
        )
    return max(values[0], values[2])


def get_rutting_strain(file_path):
    num_pattern = r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[Ee][+-]?\d+)?"
    values = []
    with open(file_path, "r", encoding="utf-8", errors="replace") as file:
        start_parsing = False
        for line in file:
            if "epZ" in line:
                start_parsing = True
                continue
            if start_parsing:
                numbers = re.findall(num_pattern, line)
                if len(numbers) >= 8:
                    values.append(abs(float(numbers[7])))

    if len(values) < 8:
        raise ValueError(
            f"Expected at least 8 epZ strain readings in {file_path}; found {len(values)}"
        )
    return max(values[5], values[7])


def main():
    jar_directory = str(find_iitpave_directory(SCRIPT_DIR))
    global JAR_DIRECTORY, JAR_NAME, INPUT_FILE_NAME, OUTPUT_FILE_NAME
    JAR_DIRECTORY = jar_directory
    JAR_NAME = "IITPFILE.exe"
    INPUT_FILE_NAME = "IITPAVE.IN"
    OUTPUT_FILE_NAME = "IITPAVE.out"

    workbook_path = SCRIPT_DIR / "Modulus_Calculation.xlsx"
    LOGGER.info(
        "Starting remaining-life calculation; workbook=%s; IITPAVE directory=%s",
        workbook_path,
        JAR_DIRECTORY,
    )

    with xw.App(visible=False) as app:
        wb3 = app.books.open(str(workbook_path))
        try:
            ws3 = wb3.sheets[0]
            LOGGER.info("Opened workbook sheet %s", ws3.name)
            if wb3.api.ReadOnly:
                raise PermissionError(
                    "Modulus_Calculation.xlsx is open read-only; "
                    "close it in Excel/WPS and retry."
                )

            # The inserted IITPAVE deflection columns X:AF shift all former
            # columns X onward nine places to the right.
            row_number = 3
            while ws3.range(f"A{row_number}").value is not None:
                temperature = ws3.range(f"AH{row_number}").value
                if not isinstance(temperature, (int, float)) or temperature <= 0:
                    raise ValueError(
                        f"Invalid asphalt temperature in AH{row_number}: {temperature!r}"
                    )

                ws3.range(f"AI{row_number}").value = (
                    (1 - 0.238 * math.log(35))
                    / (1 - 0.238 * math.log(temperature))
                )
                ws3.range(f"AK{row_number}").value = (
                    ws3.range(f"AI{row_number}").value
                    * ws3.range(f"U{row_number}").value
                )

                season = ws3.range(f"AJ{row_number}").value
                base_modulus = ws3.range(f"V{row_number}").value
                subgrade_modulus = ws3.range(f"W{row_number}").value

                if season == "Summer":
                    ws3.range(f"AL{row_number}").value = (
                        -0.0003 * (base_modulus ** 2)
                        + 0.9854 * base_modulus
                        - 32.989
                    )
                    ws3.range(f"AM{row_number}").value = (
                        0.8554 * subgrade_modulus - 8.461
                    )
                elif season == "Winter":
                    ws3.range(f"AL{row_number}").value = (
                        10.5523 * (base_modulus ** 0.624) - 113.857
                    )
                    ws3.range(f"AM{row_number}").value = (
                        3.351 * (subgrade_modulus ** 0.7688) - 28.9
                    )
                else:
                    ws3.range(f"AL{row_number}").value = base_modulus
                    ws3.range(f"AM{row_number}").value = subgrade_modulus

                row_number += 1

            LOGGER.info(
                "Temperature and seasonal corrections completed for %d rows",
                row_number - 3,
            )

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
                    numeric_boundaries = [float(token) for token in tokens[:-1]]
                    if any(not math.isfinite(value) for value in numeric_boundaries):
                        raise ValueError("All numeric boundaries must be finite numbers.")
                    if any(
                        upper <= lower
                        for lower, upper in zip(
                            numeric_boundaries, numeric_boundaries[1:]
                        )
                    ):
                        raise ValueError(
                            "Boundaries must be strictly increasing with no duplicates."
                        )
                    break
                except ValueError as exc:
                    print(f"Invalid boundaries: {exc}")

            # Chainage values in A are interpreted as kilometres.
            data_rows = []
            row_number = 3
            while ws3.range(f"A{row_number}").value is not None:
                raw_chainage = ws3.range(f"A{row_number}").value
                try:
                    chainage_km = float(raw_chainage)
                except (TypeError, ValueError):
                    raise ValueError(
                        f"Invalid/non-numeric chainage in A{row_number}: {raw_chainage!r}"
                    )
                if not math.isfinite(chainage_km):
                    raise ValueError(
                        f"Non-finite chainage in A{row_number}: {raw_chainage!r}"
                    )
                data_rows.append((row_number, chainage_km))
                row_number += 1

            if not data_rows:
                raise ValueError("No data rows found in column A starting at row 3.")

            max_chainage_km = max(chainage for _, chainage in data_rows)
            if max_chainage_km <= numeric_boundaries[-1]:
                raise ValueError(
                    f"The last numeric boundary ({numeric_boundaries[-1]:.6g} km) must be "
                    f"less than the maximum chainage ({max_chainage_km:.6g} km), "
                    "because 'end' resolves to the maximum chainage in column A."
                )

            boundaries = numeric_boundaries + [max_chainage_km]
            intervals = []
            for lower_km, upper_km in zip(boundaries, boundaries[1:]):
                section_rows = []
                for excel_row, chainage_km in data_rows:
                    if lower_km == 0:
                        belongs = 0 <= chainage_km <= upper_km
                    else:
                        belongs = lower_km < chainage_km <= upper_km
                    if belongs:
                        section_rows.append((excel_row, chainage_km))
                if section_rows:
                    intervals.append((lower_km, upper_km, section_rows))
                else:
                    LOGGER.warning(
                        "No readings found in requested interval %.6g-%.6g km; skipping.",
                        lower_km,
                        upper_km,
                    )

            # Section outputs were formerly AE:AP; they now begin at AN and end at AY.
            last_data_row = max(row for row, _ in data_rows)
            ws3.range(f"AN3:AY{last_data_row}").clear_contents()

            LOGGER.info(
                "Using variable chainage boundaries in km: %s; end=%.6g km; "
                "%d rows across %d populated sections",
                ", ".join(f"{value:g}" for value in numeric_boundaries),
                max_chainage_km,
                len(data_rows),
                len(intervals),
            )

            for section_number, (lower_km, upper_km, section_rows) in enumerate(
                intervals, start=1
            ):
                start_row = min(row for row, _ in section_rows)
                end_row = max(row for row, _ in section_rows)
                start_chainage = min(value for _, value in section_rows)
                end_chainage = max(value for _, value in section_rows)

                thickness = [
                    ws3.range(f"M{start_row}").value,
                    ws3.range(f"N{start_row}").value,
                ]
                thickness_pairs = {
                    (ws3.range(f"M{row}").value, ws3.range(f"N{row}").value)
                    for row, _ in section_rows
                }
                if len(thickness_pairs) > 1:
                    LOGGER.warning(
                        "Interval %d (rows %d-%d; %.6g-%.6g km) contains varying "
                        "M/N thicknesses %r. IITPAVE will use the first row's thicknesses %r.",
                        section_number,
                        start_row,
                        end_row,
                        start_chainage,
                        end_chainage,
                        sorted(thickness_pairs, key=str),
                        thickness,
                    )

                bituminous_values = [
                    ws3.range(f"AK{row}").value for row, _ in section_rows
                ]
                base_values = [
                    ws3.range(f"AL{row}").value for row, _ in section_rows
                ]
                subgrade_values = [
                    ws3.range(f"AM{row}").value for row, _ in section_rows
                ]

                bituminous_values = [
                    float(value) for value in bituminous_values if value is not None
                ]
                base_values = [float(value) for value in base_values if value is not None]
                subgrade_values = [
                    float(value) for value in subgrade_values if value is not None
                ]
                if not bituminous_values or not base_values or not subgrade_values:
                    raise ValueError(
                        f"Interval {section_number} (rows {start_row}-{end_row}) has "
                        "missing corrected modulus values in AK:AM."
                    )

                # Section-level outputs are written at the first row of each section.
                ws3.range(f"AN{start_row}").value = section_number
                ws3.range(f"AO{start_row}").value = lower_km
                ws3.range(f"AP{start_row}").value = upper_km
                ws3.range(f"AQ{start_row}").value = upper_km - lower_km
                ws3.range(f"AR{start_row}").value = np.quantile(bituminous_values, 0.15)
                ws3.range(f"AS{start_row}").value = np.quantile(base_values, 0.15)
                ws3.range(f"AT{start_row}").value = np.quantile(subgrade_values, 0.15)

                moduli = [
                    ws3.range(f"AR{start_row}").value,
                    ws3.range(f"AS{start_row}").value,
                    ws3.range(f"AT{start_row}").value,
                ]
                number_of_layers = len(moduli)
                poissons_ratio = np.full(3, 0.4)
                poissons_ratio[0] = 0.5

                output_path = run_iitpave_jar(
                    moduli, thickness, number_of_layers, poissons_ratio
                )
                fatigue_strain = get_fatigue_strain(output_path)
                rutting_strain = get_rutting_strain(output_path)
                ws3.range(f"AU{start_row}").value = fatigue_strain
                ws3.range(f"AV{start_row}").value = rutting_strain

                bituminous_e = float(ws3.range(f"AR{start_row}").value)
                if bituminous_e <= 0 or fatigue_strain <= 0 or rutting_strain <= 0:
                    raise ValueError(
                        f"Interval {section_number}: modulus and strains must be positive; "
                        f"AR={bituminous_e}, AU={fatigue_strain}, AV={rutting_strain}."
                    )

                ws3.range(f"AW{start_row}").value = (
                    (0.711 * 10**-4)
                    * ((1 / fatigue_strain) ** 3.89)
                    * ((1 / bituminous_e) ** 0.854)
                    * (10**-6)
                )
                ws3.range(f"AX{start_row}").value = (
                    (1.41 * 10**-8)
                    * ((1 / rutting_strain) ** 4.5337)
                    * (10**-6)
                )
                ws3.range(f"AY{start_row}").value = min(
                    ws3.range(f"AW{start_row}").value,
                    ws3.range(f"AX{start_row}").value,
                )

                LOGGER.info(
                    "Section %d completed: rows %d-%d, %.6g-%.6g km; "
                    "AR=%r, AU=%r, AV=%r",
                    section_number,
                    start_row,
                    end_row,
                    lower_km,
                    upper_km,
                    bituminous_e,
                    fatigue_strain,
                    rutting_strain,
                )

            wb3.save()
            LOGGER.info(
                "Remaining-life calculation completed; populated sections=%d",
                len(intervals),
            )
        finally:
            wb3.close()


if __name__ == "__main__":
    main()
