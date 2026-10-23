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

# IITPAVE configuration
JAR_DIRECTORY = str(find_iitpave_directory(SCRIPT_DIR))
JAR_NAME = "IITPFILE.exe"
INPUT_FILE_NAME = "IITPAVE.IN"
OUTPUT_FILE_NAME = "IITPAVE.out"

jar_path = os.path.join(JAR_DIRECTORY, JAR_NAME)
input_path = os.path.join(JAR_DIRECTORY, INPUT_FILE_NAME)
output_path = os.path.join(JAR_DIRECTORY, OUTPUT_FILE_NAME)


def run_iitpave_jar(moduli, thicknesses, number, poi):
    """Write IITPAVE input and execute IITPAVE."""
    if number == 3:
        input_content = f"""{number}
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
        input_content = f"""{number}
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

    with open(input_path, "w", encoding="utf-8") as input_file:
        input_file.write(input_content)

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


def _get_strain_values(file_path, header, value_index):
    """Extract the requested strain values from IITPAVE output."""
    num_pattern = r"[-+]?(?:\d+\.\d*|\.\d+|\d+)(?:[Ee][-+]?\d+)?"
    values = []
    parsing = False

    with open(file_path, "r", encoding="utf-8", errors="replace") as output_file:
        for line in output_file:
            if header in line:
                parsing = True
                continue

            if parsing:
                numbers = re.findall(num_pattern, line)
                if len(numbers) > value_index:
                    try:
                        values.append(abs(float(numbers[value_index])))
                    except ValueError:
                        continue

    return values


def get_fatigue_strain(file_path):
    values = _get_strain_values(file_path, "epT", 8)
    if len(values) < 3:
        raise ValueError(
            f"Expected at least 3 fatigue-strain records in {file_path}; "
            f"found {len(values)}."
        )
    return max(values[0], values[2])


def get_rutting_strain(file_path):
    values = _get_strain_values(file_path, "epZ", 7)
    if len(values) < 8:
        raise ValueError(
            f"Expected at least 8 rutting-strain records in {file_path}; "
            f"found {len(values)}."
        )
    return max(values[5], values[7])


def main():
    LOGGER.info(
        "Starting overlay assessment; workbook=%s; IITPAVE directory=%s",
        SCRIPT_DIR / "Modulus_Calculation.xlsx",
        JAR_DIRECTORY,
    )

    with xw.App(visible=False) as app:
        wb = app.books.open(str(SCRIPT_DIR / "Modulus_Calculation.xlsx"))
        try:
            ws = wb.sheets[0]
            LOGGER.info("Opened workbook sheet %s", ws.name)

            if wb.api.ReadOnly:
                raise PermissionError(
                    "Modulus_Calculation.xlsx is open read-only; "
                    "close it in Excel/WPS and retry."
                )

            msa_required1 = 500
            msa_required2 = 500
            moduli_new = 3000
            poi_new = 0.35

            row = 3
            processed_rows = 0
            skipped_rows = 0

            while ws.range(f"A{row}").value is not None and row < 1000:
                chainage = ws.range(f"A{row}").value
                remaining_life = ws.range(f"AY{row}").value

                if remaining_life is None:
                    LOGGER.warning(
                        "Skipping row %d (chainage=%s): remaining life AY is blank; "
                        "run REMAINING_LIFE.py first",
                        row,
                        chainage,
                    )
                    skipped_rows += 1
                    row += 1
                    continue

                try:
                    chainage_value = float(chainage)
                    remaining_life_value = float(remaining_life)
                except (TypeError, ValueError):
                    LOGGER.warning(
                        "Skipping row %d: invalid chainage or remaining life "
                        "(chainage=%s, remaining life=%s)",
                        row,
                        chainage,
                        remaining_life,
                    )
                    skipped_rows += 1
                    row += 1
                    continue

                if remaining_life_value <= 0:
                    LOGGER.warning(
                        "Skipping row %d (chainage=%s): remaining life must be positive",
                        row,
                        chainage,
                    )
                    skipped_rows += 1
                    row += 1
                    continue

                if chainage_value < 49100:
                    msa_required = msa_required1
                elif chainage_value > 49100:
                    msa_required = msa_required2
                else:
                    # The original script has no branch for chainage exactly 49100.
                    LOGGER.warning(
                        "Skipping row %d: chainage is exactly 49100 and the original "
                        "script defines no threshold branch for it",
                        row,
                    )
                    skipped_rows += 1
                    row += 1
                    continue

                if remaining_life_value >= msa_required:
                    ws.range(f"AZ{row}").value = "No"
                    processed_rows += 1
                    LOGGER.info(
                        "Overlay assessment row %d complete: chainage=%s, overlay=No",
                        row,
                        chainage,
                    )
                    row += 1
                    continue

                ws.range(f"AZ{row}").value = "Yes"

                # Corrected field moduli are stored in AK:AM in the revised workbook.
                # These are the three existing pavement layers; the overlay is layer 1.
                moduli = [
                    moduli_new,
                    ws.range(f"AK{row}").value,
                    ws.range(f"AL{row}").value,
                    ws.range(f"AM{row}").value,
                ]
                poiss = [poi_new, 0.5, 0.4, 0.4]

                asphalt_thickness = ws.range(f"M{row}").value
                base_thickness = ws.range(f"N{row}").value
                if asphalt_thickness is None or base_thickness is None:
                    raise ValueError(
                        f"Missing thickness value in M{row} or N{row}."
                    )

                thick_req = 0
                fat_req = 0
                rut_req = 0
                fatl_req = 0
                rutl_req = 0

                # Test overlay thicknesses from 5 mm through 995 mm.
                for overlay_thickness in range(5, 1000, 5):
                    thickness = [
                        overlay_thickness,
                        float(asphalt_thickness),
                        float(base_thickness),
                    ]
                    run_iitpave_jar(moduli, thickness, 4, poiss)
                    fs = get_fatigue_strain(output_path)
                    rs = get_rutting_strain(output_path)

                    # The original calculation uses the value in AI as the
                    # modulus in the fatigue-life equation. AI is Lambda in the
                    # revised layout, so use the corrected asphalt modulus AK.
                    asphalt_modulus = float(ws.range(f"AK{row}").value)
                    if fs <= 0 or rs <= 0 or asphalt_modulus <= 0:
                        raise ValueError(
                            f"Non-positive strain/modulus at row {row}: "
                            f"fatigue={fs}, rutting={rs}, modulus={asphalt_modulus}."
                        )

                    fatigue_life = (
                        (0.711 * 10**-4)
                        * ((1 / fs) ** 3.89)
                        * ((1 / asphalt_modulus) ** 0.854)
                    ) * (10**-6)
                    rutting_life = (
                        (1.41 * 10**-8) * ((1 / rs) ** 4.5337)
                    ) * (10**-6)
                    residual_life = min(fatigue_life, rutting_life)

                    if residual_life >= msa_required:
                        thick_req = overlay_thickness
                        fat_req = fs
                        rut_req = rs
                        fatl_req = fatigue_life
                        rutl_req = rutting_life
                        break

                # Revised workbook output columns for overlay assessment.
                ws.range(f"BA{row}").value = thick_req
                ws.range(f"BB{row}").value = moduli_new
                ws.range(f"BC{row}").value = poi_new
                ws.range(f"BD{row}").value = fat_req
                ws.range(f"BE{row}").value = rut_req
                ws.range(f"BF{row}").value = fatl_req
                ws.range(f"BG{row}").value = rutl_req
                ws.range(f"BH{row}").value = min(fatl_req, rutl_req)

                processed_rows += 1
                LOGGER.info(
                    "Overlay assessment row %d complete: chainage=%s, overlay=%s, "
                    "thickness=%s",
                    row,
                    chainage,
                    ws.range(f"AZ{row}").value,
                    ws.range(f"BA{row}").value,
                )
                row += 1

            wb.save()
            LOGGER.info(
                "Overlay assessment completed; processed rows=%d, skipped rows=%d",
                processed_rows,
                skipped_rows,
            )
        finally:
            wb.close()


if __name__ == "__main__":
    main()
