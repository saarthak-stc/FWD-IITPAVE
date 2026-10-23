import argparse
import logging
import sys
from contextlib import contextmanager
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent

try:  # Use the shared pipeline logger when it is available.
    from pipeline_logging import LOG_PATH, get_logger

    LOGGER = get_logger(__file__)
except ImportError:
    LOG_PATH = None
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    LOGGER = logging.getLogger("input")

# --------------------------------------------------------------------------- #
# Defaults (unchanged from the text-file based script)
# --------------------------------------------------------------------------- #
DEFAULT_INPUT_FILE = "NHAI remuli lha3000 200 500.txt"
DEFAULT_WORKBOOK = "Modulus_Calculation.xlsx"
DEFAULT_SHEET = "RHS"
INPUT_ENCODING = "utf-16"

LOAD_N = 40000                      # Load(N) written to column B
STRESS_MPA = 0.56                   # Stress(MPa) written to column C
NORMALIZING_LOAD_KN = 40            # deflections are normalised to this load
USE_FILE_THICKNESS = True           # True: M/N come from the .txt "Layer Thickness (mm)" columns
THICKNESS = [115, 450]              # fallback bituminous, base (mm), used only if USE_FILE_THICKNESS is False
E_RANGES = [400, 3000, 100, 500, 20, 200]   # bit lo/hi, base lo/hi, subgrade lo/hi (MPa)
SEASON = "Monsoon"
MAX_DROPS_PER_POINT = 3             # drops averaged per chainage (original behaviour)

FIRST_DATA_ROW = 3
LAST_COLUMN = "BH"
DATA_WIDTH = 20                     # columns A:T are written by this script

MIN_COLUMNS = 22
METADATA_PREFIXES = (
    "Date Created", "Project Name", "Project Location", "Performer Name",
    "Client Name", "Location", "Disttrict", "Test Start", "Test End",
    "Deflection Sensors", "Date\tTime",
)

# --------------------------------------------------------------------------- #
# Worksheet layout
# --------------------------------------------------------------------------- #
MERGED_HEADERS = {
    "A1:A2": "Chainage",
    "B1:B2": "Load(N)",
    "C1:C2": "Stress(MPa)",
    "D1:L1": "Measured Deflections at various sensors",
    "M1:N1": "Thickness of each layers",
    "O1:T1": "Range of E values",
    "U1:W1": "Predicted E values",
    "X1:AF1": "IITPAVE Deflection at various sensors",
    "AG1:AG2": "RMSE Using IITPAVE(%)",
    "AH1:AH2": "Asphalt Temperature",
    "AI1:AI2": "Lambda",
    "AJ1:AJ2": "Season",
    "AK1:AM1": "Field Modulus After Temp & Seasonal Corrections",
    "AN1:AN2": "Sl No.",
    "AO1:AP1": "Chainage",
    "AQ1:AQ2": "Length (Km)",
    "AR1:AT1": "15th Percentile E Value",
    "AU1:AU2": "Fatigue Strain",
    "AV1:AV2": "Rutting Strain",
    "AW1:AW2": "Fatigue Residual Life(MSA)",
    "AX1:AX2": "Rutting Residual Life(MSA)",
    "AY1:AY2": "Remaining life(MSA) as per 90% Reliability",
    "AZ1:AZ2": "Overlay Required (Yes/No)",
    "BA1:BA2": "Overlay Thickness Required (mm)",
    "BB1:BB2": "E moduli of Overlay",
    "BC1:BC2": "Poisson's Ratio of Overlay",
    "BD1:BD2": "Fatigue Strain Post Overlay",
    "BE1:BE2": "Rutting Strain Post Overlay",
    "BF1:BF2": "Fatigue Residual Life(MSA) Post Overlay",
    "BG1:BG2": "Rutting Residual Life(MSA) Post Overlay",
    "BH1:BH2": "Remaining life(MSA) as per 90% Reliability Post Overlay",
}

SUBHEADERS = {
    "D2": "D0 (mm)\n0 mm",
    "E2": "D1 (mm)\n200 mm",
    "F2": "D2 (mm)\n300 mm",
    "G2": "D3 (mm)\n450 mm",
    "H2": "D4 (mm)\n600 mm",
    "I2": "D5 (mm)\n900 mm",
    "J2": "D6 (mm)\n1200 mm",
    "K2": "D7 (mm)\n1500 mm",
    "L2": "D8 (mm)\n1800 mm",
    "M2": "Bituminous Layer Thickness",
    "N2": "Base Layer Thickness",
    "O2": "Bit. E lower Range",
    "P2": "Bit. E upper Range",
    "Q2": "Base E lower Range",
    "R2": "Base E upper Range",
    "S2": "Sub-Grade E lower Range",
    "T2": "Sub-Grade E upper Range",
    "U2": "Bituminous (E1)",
    "V2": "Base (E2)",
    "W2": "Sub-Grade (E3)",
    "X2": "IITPD0",
    "Y2": "IITPD1",
    "Z2": "IITPD2",
    "AA2": "IITPD3",
    "AB2": "IITPD4",
    "AC2": "IITPD5",
    "AD2": "IITPD6",
    "AE2": "IITPD7",
    "AF2": "IITPD8",
    "AK2": "Bituminous (E1)",
    "AL2": "Base (E2)",
    "AM2": "Sub-Grade (E3)",
    "AO2": "From",
    "AP2": "To",
    "AR2": "Bituminous Layer",
    "AS2": "Base Layer",
    "AT2": "Sub-Grade Layer",
}

TWO_DECIMAL_COLUMNS = (
    "U", "V", "W",              # predicted E values
    "AG",                       # RMSE
    "AK", "AL", "AM",           # field modulus after corrections
    "AR", "AS", "AT",           # 15th percentile E values
    "AW", "AX", "AY",           # residual lives
    "BF", "BG", "BH",           # residual lives post overlay
)
FOUR_DECIMAL_COLUMNS = ("AI",)  # Lambda


# --------------------------------------------------------------------------- #
# Input parsing
# --------------------------------------------------------------------------- #
def parse_chainage(text):
    """Convert a chainage such as ``'0+100'`` or ``'14+690'`` to metres.

    Plain numbers (``'250'``, ``'250.5'``) are returned as floats.
    Raises ``ValueError`` for anything that cannot be parsed.
    """
    text = text.strip()
    if "+" in text:
        km_part, metre_part = text.split("+", 1)
        return int(km_part) * 1000 + float(metre_part)
    return float(text)


def get_fwd_data(input_file):
    """Read the tab-delimited FWD export and return a list of measurement rows.

    Row layout:
        0 chainage (m)      1 drop number     2 load (kN)
        3-11 D1..D9 (micron, i.e. sensors D0..D8)
        12 air temp (C)     13 surface temp (C)   14 comment
        15 bituminous thickness (mm)  16 base thickness (mm)  17 subgrade thickness (mm)
    """
    input_path = Path(input_file)
    if not input_path.is_file():
        raise FileNotFoundError(f"Input file not found: {input_path.resolve()}")

    extracted_data = []
    skipped_lines = 0
    zero_load_lines = 0

    try:
        with open(input_path, "r", encoding=INPUT_ENCODING, errors="ignore") as handle:
            for line_number, line in enumerate(handle, start=1):
                line = line.strip()
                if not line:
                    continue
                if any(line.startswith(prefix) for prefix in METADATA_PREFIXES):
                    continue

                parts = line.split("\t")
                if len(parts) < MIN_COLUMNS:
                    skipped_lines += 1
                    LOGGER.debug("Skipping short/non-data line %d", line_number)
                    continue

                try:
                    load = float(parts[6])
                    if load <= 0:
                        zero_load_lines += 1
                        LOGGER.warning(
                            "Line %d (chainage %s, drop %s): load is %s kN; row skipped",
                            line_number, parts[4], parts[5], parts[6],
                        )
                        continue
                    row = [
                        parse_chainage(parts[4]),   # 0: chainage (m)
                        int(parts[5]),              # 1: drop number
                        load,                       # 2: load (kN)
                        float(parts[7]),            # 3: D1 (centre, D0 in sheet)
                        float(parts[8]),            # 4
                        float(parts[9]),            # 5
                        float(parts[10]),           # 6
                        float(parts[11]),           # 7
                        float(parts[12]),           # 8
                        float(parts[13]),           # 9
                        float(parts[14]),           # 10
                        float(parts[15]),           # 11: D9 (D8 in sheet)
                        float(parts[16]),           # 12: air temp
                        float(parts[17]),           # 13: surface temp
                        parts[18],                  # 14: comment
                        float(parts[19]),           # 15: bituminous thickness
                        float(parts[20]),           # 16: base thickness
                        float(parts[21]),           # 17: subgrade thickness
                    ]
                except ValueError:
                    skipped_lines += 1
                    LOGGER.debug("Skipping unparsable line %d: %r", line_number, line[:80])
                    continue
                extracted_data.append(row)
    except PermissionError as exc:
        raise PermissionError(f"Cannot read input file (permission denied or file locked): {input_path.resolve()}") from exc
    except OSError as exc:
        raise OSError(f"Cannot read input file {input_path.resolve()}: {exc}") from exc

    LOGGER.info(
        "Read %d valid measurement rows from %s; skipped %d non-data/invalid lines and %d zero-load rows",
        len(extracted_data), input_path.resolve(), skipped_lines, zero_load_lines,
    )
    if not extracted_data:
        raise ValueError(
            f"No valid FWD measurement rows found in {input_path.resolve()}. "
            "Expected a tab-delimited UTF-16 export with at least 22 columns per data row."
        )

    return extracted_data


# --------------------------------------------------------------------------- #
# Row preparation
# --------------------------------------------------------------------------- #
def group_by_chainage(result, max_group=MAX_DROPS_PER_POINT):
    """Group consecutive records sharing a chainage (at most ``max_group`` per group)."""
    groups = []
    index = 0
    while index < len(result):
        chainage = result[index][0]
        start = index
        while (
            index < len(result)
            and result[index][0] == chainage
            and index - start < max_group
        ):
            index += 1
        if index < len(result) and result[index][0] == chainage:
            LOGGER.warning(
                "Chainage %s has more than %d consecutive drops; the extra drops form a separate row",
                chainage, max_group,
            )
        groups.append(result[start:index])
    return groups


def _clean_number(value):
    """200.0 -> 200 so whole-number thicknesses appear as integers in Excel."""
    return int(value) if float(value).is_integer() else value


def create_rows(result):
    """Build one output row per chainage group.

    Returns a list of ``(values_A_to_T, asphalt_temperature, drops_in_group)``.
    Deflections are normalised to 40 kN, averaged over the group and converted
    from micron to mm (``/1000``) - identical to the original script.
    """
    rows = []
    for group in group_by_chainage(result):
        count = len(group)
        deflections = []
        for sensor in range(9):
            total = sum(
                (record[sensor + 3] / record[2]) * NORMALIZING_LOAD_KN for record in group
            )
            deflections.append(total / (count * 1000))
        if USE_FILE_THICKNESS:
            thickness = [_clean_number(group[0][15]), _clean_number(group[0][16])]
            if any((r[15], r[16]) != (group[0][15], group[0][16]) for r in group):
                LOGGER.warning("Chainage %s: drops list different layer thicknesses; using the first drop", group[0][0])
        else:
            thickness = list(THICKNESS)
        values = [
            group[0][0],            # A: chainage (m)
            LOAD_N,                 # B
            STRESS_MPA,             # C
            *deflections,           # D:L
            thickness[0],           # M: bituminous thickness (mm)
            thickness[1],           # N: base thickness (mm)
            *E_RANGES,              # O:T
        ]
        rows.append((values, group[0][13], count))   # asphalt temp = surface temp (as before)
    return rows


# --------------------------------------------------------------------------- #
# Excel helpers
# --------------------------------------------------------------------------- #
@contextmanager
def excel_quiet(app):
    """Disable screen updating, alerts and events; always restore previous state."""
    previous = (app.screen_updating, app.display_alerts, app.enable_events)
    app.screen_updating = False
    app.display_alerts = False
    app.enable_events = False
    try:
        yield
    finally:
        try:
            app.screen_updating, app.display_alerts, app.enable_events = previous
        except Exception:  # Excel may already be gone after a crash
            LOGGER.warning("Could not restore Excel application settings", exc_info=True)


def merger(s1, s2, sheet):
    """Unmerge, clear and merge the range ``s1:s2``."""
    cell_range = sheet.range(f"{s1}:{s2}")
    cell_range.api.UnMerge()
    cell_range.clear_contents()
    cell_range.api.Merge()


def write_headers(sheet):
    """Reset rows 1-2 (A:BH) and write all merged headers, subheaders and formatting."""
    header_area = sheet.range(f"A1:{LAST_COLUMN}2")
    header_area.api.UnMerge()
    header_area.clear_contents()

    for cell_range, title in MERGED_HEADERS.items():
        start, end = cell_range.split(":")
        merger(start, end, sheet)
        sheet.range(start).value = title

    for cell, value in SUBHEADERS.items():
        sheet.range(cell).value = value

    sheet.range(f"A1:{LAST_COLUMN}2").api.WrapText = True
    sheet.range(f"A:{LAST_COLUMN}").api.HorizontalAlignment = -4108
    sheet.range(f"A:{LAST_COLUMN}").api.VerticalAlignment = -4108
    sheet.range(f"A1:{LAST_COLUMN}2").api.Font.Bold = True
    sheet.range(f"A1:{LAST_COLUMN}2").color = (192, 192, 192)
    sheet.range(f"A:{LAST_COLUMN}").api.Borders.LineStyle = 1
    for column in TWO_DECIMAL_COLUMNS:
        sheet.range(f"{column}:{column}").number_format = "0.00"
    for column in FOUR_DECIMAL_COLUMNS:
        sheet.range(f"{column}:{column}").number_format = "0.0000"


def clear_old_rows(sheet):
    """Clear stale values below the headers (A:BH only, this sheet only)."""
    last_row = sheet.used_range.last_cell.row
    if last_row >= FIRST_DATA_ROW:
        sheet.range(f"A{FIRST_DATA_ROW}:{LAST_COLUMN}{last_row}").clear_contents()
        LOGGER.info("Cleared old data in A%d:%s%d", FIRST_DATA_ROW, LAST_COLUMN, last_row)


def write_rows(sheet, rows):
    """Write imported data. X:AG, AI and AK:BH are left blank for later stages."""
    first = FIRST_DATA_ROW
    last = first + len(rows) - 1
    for values, _, _ in rows:
        if len(values) != DATA_WIDTH:
            raise ValueError(f"Internal error: expected {DATA_WIDTH} values per row, got {len(values)}")
    sheet.range(f"A{first}:T{last}").value = [values for values, _, _ in rows]
    sheet.range(f"AH{first}:AH{last}").value = [[temperature] for _, temperature, _ in rows]
    sheet.range(f"AJ{first}:AJ{last}").value = [[SEASON] for _ in rows]


def resolve_path(path):
    """Use the path as given (relative to the working directory); if a relative
    file is missing there but present next to this script, use that one."""
    path = Path(path)
    if not path.is_absolute() and not path.exists() and (SCRIPT_DIR / path).exists():
        path = SCRIPT_DIR / path
    return path.resolve()


# --------------------------------------------------------------------------- #
# Main workflow
# --------------------------------------------------------------------------- #
def run(input_file, workbook_path, sheet_name=DEFAULT_SHEET):
    input_path = resolve_path(input_file)
    workbook_path = resolve_path(workbook_path)

    result = get_fwd_data(input_path)
    rows = create_rows(result)
    LOGGER.info("Prepared %d chainage rows", len(rows))

    try:
        import xlwings as xw
    except ImportError as exc:
        raise RuntimeError("xlwings is not installed (pip install xlwings); Microsoft Excel is also required") from exc

    try:
        with xw.App(visible=False) as app:
            with excel_quiet(app):
                if workbook_path.exists():
                    workbook = app.books.open(str(workbook_path))
                    LOGGER.info("Opened workbook: %s", workbook_path)
                else:
                    workbook = app.books.add()
                    workbook.save(str(workbook_path))
                    LOGGER.info("Created workbook: %s", workbook_path)
                try:
                    if getattr(workbook.api, "ReadOnly", False):
                        raise PermissionError(
                            f"Workbook is open read-only: {workbook_path}. Close it in Excel "
                            "(and remove any stale ~$ lock file) before retrying."
                        )
                    if sheet_name in [sheet.name for sheet in workbook.sheets]:
                        sheet = workbook.sheets[sheet_name]
                    else:
                        sheet = workbook.sheets.add(sheet_name, before=workbook.sheets[0])
                        LOGGER.info("Created worksheet: %s", sheet_name)

                    write_headers(sheet)
                    clear_old_rows(sheet)
                    write_rows(sheet, rows)
                    for row_number, (values, _, count) in enumerate(rows, start=FIRST_DATA_ROW):
                        LOGGER.info(
                            "Point %d done: row %d, chainage=%s, grouped drops=%d",
                            row_number - FIRST_DATA_ROW + 1, row_number, values[0], count,
                        )
                    workbook.save()
                    LOGGER.info("Saved %d rows to %s (sheet %s)", len(rows), workbook_path, sheet_name)
                finally:
                    workbook.close()
    except (PermissionError, FileNotFoundError, ValueError):
        raise
    except Exception as exc:
        raise RuntimeError(f"Excel automation failed while updating {workbook_path}: {exc}") from exc


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Import FWD deflection measurements (.txt export) into the Modulus Calculation workbook."
    )
    parser.add_argument("--input", default=DEFAULT_INPUT_FILE, help=f"Input .txt file (default: {DEFAULT_INPUT_FILE})")
    parser.add_argument("--workbook", default=DEFAULT_WORKBOOK, help=f"Output workbook (default: {DEFAULT_WORKBOOK})")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help=f"Target worksheet (default: {DEFAULT_SHEET})")
    args = parser.parse_args(argv)

    LOGGER.info("Starting FWD import; common log: %s", LOG_PATH)
    try:
        run(args.input, args.workbook, args.sheet)
    except Exception as exc:
        LOGGER.exception("FWD import failed: %s", exc)
        return 1
    LOGGER.info("FWD import completed successfully")
    return 0


if __name__ == "__main__":
    sys.exit(main())