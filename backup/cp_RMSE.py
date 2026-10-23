import numpy as np
from scipy.optimize import minimize  # Kept because this module may use it in related workflows.
import subprocess
import os
import xlwings as xw
import re

from pathlib import Path
from pipeline_logging import find_iitpave_directory, get_logger

SCRIPT_DIR = Path(__file__).resolve().parent
LOGGER = get_logger(__file__)


def run_iitpave_jar(moduli, thicknesses, number, poi, pos):
    """Write IITPAVE input and run the solver."""
    jar_path = os.path.join(JAR_DIRECTORY, JAR_NAME)
    input_path = os.path.join(JAR_DIRECTORY, INPUT_FILE_NAME)
    output_path = os.path.join(JAR_DIRECTORY, OUTPUT_FILE_NAME)

    f_list = [f"0 {item}" for item in pos]
    f_list_string = "\n".join(f_list)
    input_content = (
        f"{number}\n"
        f"{' '.join(map(str, moduli))}\n"
        f"{' '.join(map(str, poi))}\n"
        f"{' '.join(map(str, thicknesses))}\n"
        "40000 0.56\n"
        f"{len(pos)}\n"
        f"{f_list_string}\n"
        "1\n"
    )

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


def get_dispz_regex(file_path):
    """Extract DispZ values from the IITPAVE output."""
    num_pattern = r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[Ee][+-]?\d+)?"
    dispz_values = []

    with open(file_path, "r", encoding="utf-8", errors="replace") as file:
        lines = file.readlines()

    start_parsing = False
    for line in lines:
        if "DispZ" in line:
            start_parsing = True
            continue

        if start_parsing:
            numbers = re.findall(num_pattern, line)
            # Original code's comment says DispZ is the 6th number (index 5),
            # but it accessed numbers[6]. Preserve that indexing to avoid
            # silently changing the solver-output interpretation.
            if len(numbers) >= 7:
                dispz_values.append(float(numbers[6]))

    return dispz_values


def RMSE_function(moduli):
    run_iitpave_jar(
        moduli, thickness, number_of_layers, poissons_ratio, postn
    )
    theoretical_deflection = np.array(get_dispz_regex(output_path), dtype=float)

    if theoretical_deflection.size != defln.size:
        raise ValueError(
            "IITPAVE returned a different number of deflections than measured: "
            f"{theoretical_deflection.size} vs {defln.size}"
        )
    if defln.size == 0:
        raise ValueError("No measured deflections are available to calculate RMSE.")

    rmse = np.sqrt(np.mean((theoretical_deflection - defln) ** 2))
    mean_deflection = np.mean(defln)
    if mean_deflection == 0:
        raise ZeroDivisionError("Mean measured deflection is zero; percentage RMSE is undefined.")
    return (rmse / mean_deflection) * 100


# --- Configuration ---
JAR_DIRECTORY = str(find_iitpave_directory(SCRIPT_DIR))
JAR_NAME = "IITPFILE.exe"
INPUT_FILE_NAME = "IITPAVE.IN"
OUTPUT_FILE_NAME = "IITPAVE.out"
jar_path = os.path.join(JAR_DIRECTORY, JAR_NAME)
input_path = os.path.join(JAR_DIRECTORY, INPUT_FILE_NAME)
output_path = os.path.join(JAR_DIRECTORY, OUTPUT_FILE_NAME)

LOGGER.info(
    "Starting RMSE; workbook=%s; IITPAVE directory=%s",
    SCRIPT_DIR / "Modulus_Calculation.xlsx",
    JAR_DIRECTORY,
)


def main():
    global postn, thickness, number_of_layers, poissons_ratio, defln

    workbook_path = SCRIPT_DIR / "Modulus_Calculation.xlsx"
    with xw.App(visible=False) as app:
        wb1 = app.books.open(str(workbook_path))
        try:
            ws1 = wb1.sheets[0]
            LOGGER.info("Opened workbook sheet %s", ws1.name)
            if wb1.api.ReadOnly:
                raise PermissionError(
                    "Modulus_Calculation.xlsx is open read-only; "
                    "close it in Excel/WPS and retry."
                )

            row_number = 3
            rel_positions = [0, 200, 300, 450, 600, 900, 1200, 1500, 1800]

            while ws1.range(f"W{row_number}").value is not None:
                measured_deflections = []
                postn = []

                # Measured deflections remain in D:L after the new X:AF section is added.
                for i in range(9):
                    sensor_value = ws1.range(
                        f"{chr(ord('D') + i)}{row_number}"
                    ).value
                    if isinstance(sensor_value, (int, float)):
                        measured_deflections.append(float(sensor_value))
                        postn.append(rel_positions[i])

                if not measured_deflections:
                    raise ValueError(
                        f"Row {row_number} has no numeric measured deflections in D:L."
                    )

                defln = np.array(measured_deflections, dtype=float)
                postn = np.array(postn, dtype=float)

                thickness = np.array(
                    [
                        ws1.range(f"M{row_number}").value,
                        ws1.range(f"N{row_number}").value,
                    ],
                    dtype=float,
                )
                number_of_layers = thickness.size + 1
                poissons_ratio = np.full(number_of_layers, 0.4)
                poissons_ratio[0] = 0.5

                moduli = np.array(
                    [
                        ws1.range(f"U{row_number}").value,
                        ws1.range(f"V{row_number}").value,
                        ws1.range(f"W{row_number}").value,
                    ],
                    dtype=float,
                )

                # RMSE moves from X to AG because X:AF is now the IITPAVE
                # deflection section. Write only the RMSE result in AG.
                rmse_value = RMSE_function(moduli)
                ws1.range(f"AG{row_number}").value = rmse_value

                LOGGER.info(
                    "Point %d RMSE=%.4f%%", row_number - 2, rmse_value
                )
                row_number += 1

            wb1.save()
            LOGGER.info("RMSE completed successfully; processed %d points", row_number - 3)
        finally:
            wb1.close()


if __name__ == "__main__":
    main()
