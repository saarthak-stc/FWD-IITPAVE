import argparse
import os
import re
import subprocess
import tempfile
import time
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait

import xlwings as xw

from pipeline_logging import get_logger

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOGGER = get_logger(__file__)

KGP_NAME = os.path.join(SCRIPT_DIR, "KGPBACK.exe")
OUTPUT_BASENAME = "backout"
# Kept for compatibility; each worker now uses its own backout file instead.
OUTPUT_FILE_NAME = os.path.join(SCRIPT_DIR, OUTPUT_BASENAME)
WORKBOOK_PATH = os.path.join(SCRIPT_DIR, "NHAI nh 49 shambhalpur rhs ch 294 276 to 268 076 Modulus_Calculation.xlsx")

FIRST_DATA_ROW = 3
LAST_INPUT_COLUMN = 20  # column T
RESULT_FIRST_COLUMN = 21  # column U (U:W receive Surface, Base, Subgrade)
REL_POSITIONS = [0, 200, 300, 450, 600, 900, 1200, 1500, 1800]
DEFAULT_MAX_WORKERS = 8


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


def run_kgpback_jar(deflections, thicknesses, modulus_ranges, positions, work_dir=None):
    """Run KGPBACK.exe for one point and return the path of its output file.

    KGPBACK.exe always writes a file called ``backout`` into its working
    directory. Passing a private ``work_dir`` lets several copies run at the
    same time without overwriting each other's output.
    """
    input_content = build_kgpback_input(
        deflections, thicknesses, modulus_ranges, positions
    )
    work_dir = work_dir or SCRIPT_DIR
    try:
        LOGGER.info("Running KGP back-calculation: %s", KGP_NAME)
        proc = subprocess.Popen(
            [KGP_NAME],
            cwd=work_dir,
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
    return os.path.join(work_dir, OUTPUT_BASENAME)


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


def read_points(ws1):
    """Read every data row once and return a list of point dicts.

    Rows are read until the first empty cell in column A, exactly like the
    original row-by-row loop. All validation happens here, before any
    back-calculation starts.
    """
    last_row = ws1.used_range.last_cell.row
    if last_row < FIRST_DATA_ROW:
        return []

    block = ws1.range(
        (FIRST_DATA_ROW, 1), (last_row, LAST_INPUT_COLUMN)
    ).options(ndim=2).value

    points = []
    for offset, cells in enumerate(block):
        chainage = cells[0]  # column A
        if chainage is None:
            break
        row_number = FIRST_DATA_ROW + offset

        deflections = []
        positions = []
        # Measured sensor deflections are in D:L (column indexes 3..11).
        for sensor_index in range(9):
            sensor_value = cells[3 + sensor_index]
            if isinstance(sensor_value, (int, float)):
                deflections.append(sensor_value)
                positions.append(REL_POSITIONS[sensor_index])

        if not 4 <= len(deflections) <= 9:
            raise ValueError(
                f"Row {row_number} "
                f"(chainage={chainage}) has "
                f"{len(deflections)} numeric sensor readings; "
                "expected between 4 and 9"
            )

        points.append(
            {
                "row": row_number,
                "chainage": chainage,
                "deflections": deflections,
                "positions": positions,
                "thicknesses": [cells[12], cells[13]],  # M, N
                "modulus_ranges": list(cells[14:20]),  # O:T
            }
        )
    return points


def solve_point(point):
    """Back-calculate one point inside its own temporary working folder."""
    LOGGER.info(
        "Processing row %d (chainage=%s) with %d sensor readings",
        point["row"],
        point["chainage"],
        len(point["deflections"]),
    )
    with tempfile.TemporaryDirectory(prefix="kgpback_") as work_dir:
        output_path = run_kgpback_jar(
            point["deflections"],
            point["thicknesses"],
            point["modulus_ranges"],
            point["positions"],
            work_dir=work_dir,
        )
        return get_dispz_regex(output_path)


def resolve_workers(workers, point_count):
    if workers is None:
        env_value = os.environ.get("KGP_WORKERS")
        if env_value:
            workers = int(env_value)
        else:
            workers = min(os.cpu_count() or 1, DEFAULT_MAX_WORKERS)
    return max(1, min(int(workers), point_count))


def main(workers=None):
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

            points = read_points(ws1)
            if not points:
                LOGGER.warning("No data rows found; nothing to do")
                return

            worker_count = resolve_workers(workers, len(points))
            LOGGER.info(
                "Back-calculating %d points with %d parallel KGPBACK worker(s)",
                len(points),
                worker_count,
            )

            started = time.perf_counter()
            results = {}  # index in points -> [surface, base, subgrade]
            failures = {}  # index in points -> exception

            with ThreadPoolExecutor(max_workers=worker_count) as pool:
                futures = {
                    pool.submit(solve_point, point): index
                    for index, point in enumerate(points)
                }
                pending = set(futures)
                while pending:
                    done, pending = wait(pending, return_when=FIRST_COMPLETED)
                    for future in done:
                        index = futures[future]
                        point = points[index]
                        try:
                            results[index] = future.result()
                        except Exception as exc:
                            failures[index] = exc
                            LOGGER.error(
                                "Point %d (row %d) failed: %s",
                                index + 1,
                                point["row"],
                                exc,
                            )
                            # Stop starting new points; points already
                            # running are allowed to finish.
                            for other in list(pending):
                                if other.cancel():
                                    pending.discard(other)
                            continue
                        LOGGER.info(
                            "Point %d completed; moduli=%s (%d/%d done)",
                            index + 1,
                            results[index],
                            len(results),
                            len(points),
                        )

            # Write U:W once. Only the unbroken run of points from the top is
            # written, matching the old row-by-row behavior where a failure
            # stopped the run after saving the rows above it.
            ordered = []
            for index in range(len(points)):
                if index not in results:
                    break
                ordered.append(results[index])

            if ordered:
                ws1.range((FIRST_DATA_ROW, RESULT_FIRST_COLUMN)).value = ordered
            wb1.save()

            LOGGER.info(
                "KGPBack processed %d of %d points in %.1f s",
                len(ordered),
                len(points),
                time.perf_counter() - started,
            )
            if failures:
                first_failed = min(failures)
                discarded = len(results) - len(ordered)
                if discarded:
                    LOGGER.warning(
                        "%d completed point(s) below the first failure were not "
                        "written; rerun after fixing the failure",
                        discarded,
                    )
                raise failures[first_failed]

            LOGGER.info(
                "KGPBack completed successfully; processed %d points", len(points)
            )
        finally:
            wb1.close()


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="KGPBACK back-calculation step")
    parser.add_argument(
        "--workers",
        type=int,
        default=None,
        help="number of KGPBACK.exe copies to run at once "
        "(default: CPU count, at most %d; env KGP_WORKERS also works)"
        % DEFAULT_MAX_WORKERS,
    )
    return parser.parse_args(argv)


if __name__ == "__main__":
    main(workers=parse_args().workers)
