import subprocess
import os
import xlwings as xw
import re
import numpy as np

from pipeline_logging import get_logger

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOGGER = get_logger(__file__)

KGP_NAME = os.path.join(SCRIPT_DIR, "KGPBACK.exe")
OUTPUT_FILE_NAME = os.path.join(SCRIPT_DIR, "backout")
WORKBOOK_PATH = os.path.join(SCRIPT_DIR, "Modulus_Calculation.xlsx")


def build_kgpback_input(deflections, thicknesses, modulus_ranges, positions):
    if len(deflections) != len(positions):
        raise ValueError(
            f"Deflection count ({len(deflections)}) does not match "
            f"sensor position count ({len(positions)})"
        )
    if not 4 <= len(deflections) <= 10:
        raise ValueError(
            f"KGPBACK requires 4 to 10 measured deflections; received {len(deflections)}"
        )
    if len(thicknesses) != 2 or len(modulus_ranges) != 6:
        raise ValueError("KGPBACK requires 2 layer thicknesses and 6 modulus range bounds")

    input_lines = [
        "40000",
        "0.56",
        str(len(deflections)),
        *(str(value) for value in positions),
        *(str(value) for value in deflections),
        *(str(value) for value in thicknesses),
        "0.5",
        "0.4",
        "0.4",
        *(str(value) for value in modulus_ranges),
    ]
    return "\n".join(input_lines) + "\n"


def run_kgpback_jar(deflections, thicknesses, modulus_ranges, positions):
    input_content = build_kgpback_input(
        deflections, thicknesses, modulus_ranges, positions
    )
    try:
        LOGGER.info("Running KGP back-calculation: %s", KGP_NAME)
        proc = subprocess.Popen(
            [KGP_NAME],
            cwd=SCRIPT_DIR,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        stdout, stderr = proc.communicate(input=input_content)
        if proc.returncode:
            LOGGER.error(
                "KGPBACK failed with exit code %d; stdout=%s; stderr=%s",
                proc.returncode,
                stdout,
                stderr,
            )
            raise subprocess.CalledProcessError(
                proc.returncode, KGP_NAME, output=stdout, stderr=stderr
            )
        LOGGER.info(
            "KGPBACK completed successfully with %d positions and %d deflections",
            len(positions),
            len(deflections),
        )
    except (subprocess.CalledProcessError, FileNotFoundError):
        LOGGER.exception("KGPBACK execution failed")
        raise


def get_dispz_regex(file_path):
    # Matches values such as -0.5525E+00, 200.00, and 0.00.
    num_pattern = r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[Ee][+-]?\d+)?"
    dispz_values = []

    with open(file_path, "r") as file:
        for line in file:
            if any(
                label in line
                for label in ("Surface (MPa)", "Base (MPa)", "Subgrade (MPa)")
            ):
                numbers = re.findall(num_pattern, line)
                if numbers:
                    dispz_values.append(float(numbers[0]))

    if len(dispz_values) != 3:
        raise ValueError(
            f"Expected 3 layer modulus values in {file_path}, found {len(dispz_values)}"
        )
    return dispz_values


def main():
    LOGGER.info("Starting KGPBack; workbook=%s", WORKBOOK_PATH)

    with xw.App(visible=False) as app:
        wb1 = app.books.open(WORKBOOK_PATH)
        try:
            if wb1.api.ReadOnly:
                raise PermissionError(
                    "Modulus_Calculation.xlsx is open read-only; "
                    "close it in Excel/WPS and retry."
                )

            # Keep the original worksheet-selection behavior.
            ws1 = wb1.sheets[0]
            LOGGER.info("Opened workbook sheet %s", ws1.name)

            row_number = 3
            rel_positions = [0, 200, 300, 450, 600, 900, 1200, 1500, 1800]

            while ws1.range(f"A{row_number}").value is not None:
                deflections = []
                positions = []

                # Measured sensor deflections remain in D:L.
                for sensor_index in range(9):
                    sensor_value = ws1.range(
                        f"{chr(ord('D') + sensor_index)}{row_number}"
                    ).value
                    if isinstance(sensor_value, (int, float)):
                        deflections.append(sensor_value)
                        positions.append(rel_positions[sensor_index])

                positions = np.array(positions)
                if not 4 <= len(deflections) <= 9:
                    raise ValueError(
                        f"Row {row_number} "
                        f"(chainage={ws1.range(f'A{row_number}').value}) has "
                        f"{len(deflections)} numeric sensor readings; "
                        "expected between 4 and 9"
                    )

                LOGGER.info(
                    "Processing row %d (chainage=%s) with %d sensor readings",
                    row_number,
                    ws1.range(f"A{row_number}").value,
                    len(deflections),
                )

                # These columns are all before the newly inserted X:AF section,
                # so their addresses do not change.
                thicknesses = [
                    ws1.range(f"M{row_number}").value,
                    ws1.range(f"N{row_number}").value,
                ]
                modulus_ranges = [
                    ws1.range(f"{column}{row_number}").value
                    for column in ("O", "P", "Q", "R", "S", "T")
                ]

                # Preserve the existing range values in O:T.
                for column, value in zip(
                    ("O", "P", "Q", "R", "S", "T"), modulus_ranges
                ):
                    ws1.range(f"{column}{row_number}").value = value

                run_kgpback_jar(
                    deflections, thicknesses, modulus_ranges, positions
                )
                output = get_dispz_regex(OUTPUT_FILE_NAME)

                # Predicted E values remain in U:W. The new IITPAVE
                # deflection section starts at X, so these columns do not shift.
                ws1.range(f"U{row_number}").value = output[0]
                ws1.range(f"V{row_number}").value = output[1]
                ws1.range(f"W{row_number}").value = output[2]

                wb1.save()
                LOGGER.info(
                    "Point %d completed; moduli=%s", row_number - 2, output
                )
                row_number += 1

            wb1.save()
            LOGGER.info(
                "KGPBack completed successfully; processed %d points",
                row_number - 3,
            )
        finally:
            wb1.close()


if __name__ == "__main__":
    main()
