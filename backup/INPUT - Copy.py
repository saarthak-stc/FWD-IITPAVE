import argparse
from pathlib import Path

from pipeline_logging import LOG_PATH, get_logger

SCRIPT_DIR = Path(__file__).resolve().parent
LOGGER = get_logger(__file__)


def get_fwd_data(input_file):
    input_path = Path(input_file)
    raw_data = input_path.read_bytes()
    if raw_data.startswith((b"\xff\xfe", b"\xfe\xff")):
        text = raw_data.decode("utf-16")
        encoding = "utf-16"
    else:
        try:
            text = raw_data.decode("utf-8-sig")
            encoding = "utf-8"
        except UnicodeDecodeError:
            text = raw_data.decode("utf-16")
            encoding = "utf-16"

    extracted_data = []
    skipped_lines = 0

    for line_number, line in enumerate(text.splitlines(), start=1):
        parts = line.split()
        if not parts:
            continue
        if parts[0].upper() == "D":
            parts = parts[1:]

        if len(parts) < 14:
            skipped_lines += 1
            continue

        try:
            sensor_count = 9 if len(parts) >= 16 else 7
            sensor_values = [float(value) for value in parts[3 : 3 + sensor_count]]
            sensor_values.extend([None] * (9 - sensor_count))
            metadata_index = 3 + sensor_count
            row = [
                int(float(parts[0])),
                float(parts[1]),
                float(parts[2]),
                *sensor_values,
                float(parts[metadata_index]),
                float(parts[metadata_index + 1]),
                parts[metadata_index + 2],
                parts[metadata_index + 3],
            ]
        except ValueError:
            skipped_lines += 1
            LOGGER.debug("Skipping non-data line %d: %s", line_number, line.strip())
            continue

        extracted_data.append(row)

    LOGGER.info(
        "Read %d valid measurement rows from %s (%s); skipped %d header/invalid lines",
        len(extracted_data),
        input_path.resolve(),
        encoding,
        skipped_lines,
    )
    if not extracted_data:
        raise ValueError(f"No valid FWD measurement rows found in {input_path}")
    return extracted_data


def merger(s1, s2, sheet, app):
    app.screen_updating = False
    app.display_alerts = False
    app.enable_events = False
    try:
        sheet.range(f"{s1}:{s2}").api.UnMerge()
        sheet.range(f"{s1}:{s2}").clear_contents()
        sheet.range(f"{s1}:{s2}").api.Merge()
    finally:
        app.screen_updating = True
        app.display_alerts = True
        app.enable_events = True


def write_headers(sheet, app):
    merged_headers = {
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
    for cell_range, title in merged_headers.items():
        start, end = cell_range.split(":")
        merger(start, end, sheet, app)
        sheet.range(start).value = title

    subheaders = {
        "D2": "D0 (mm)\\n0 mm",
        "E2": "D1 (mm)\\n200 mm",
        "F2": "D2 (mm)\\n300 mm",
        "G2": "D3 (mm)\\n450 mm",
        "H2": "D4 (mm)\\n600 mm",
        "I2": "D5 (mm)\\n900 mm",
        "J2": "D6 (mm)\\n1200 mm",
        "K2": "D7 (mm)\\n1500 mm",
        "L2": "D8 (mm)\\n1800 mm",
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
    for cell, value in subheaders.items():
        sheet.range(cell).value = value

    sheet.range("A1:BH2").api.WrapText = True
    sheet.range("A:BH").api.HorizontalAlignment = -4108
    sheet.range("A:BH").api.VerticalAlignment = -4108
    sheet.range("A1:BH2").api.Font.Bold = True
    sheet.range("A1:BH2").color = (192, 192, 192)
    sheet.range("A:BH").api.Borders.LineStyle = 1
    for column in ("U", "V", "W", "AG", "AK", "AL", "AM", "AR", "AS", "AT", "AW", "AX", "AY", "BF", "BG", "BH"):
        sheet.range(f"{column}:{column}").number_format = "0.00"
    sheet.range("AI:AI").number_format = "0.0000"


def create_rows(result):
    rows = []
    index = 0
    ranges = [400, 1500, 100, 500, 20, 200]
    while index < len(result):
        chainage = result[index][0]
        same_chainage = []
        while index < len(result) and result[index][0] == chainage:
            same_chainage.append(result[index])
            index += 1

        normalized_deflections = []
        for sensor in range(9):
            measurements = [
                (record[sensor + 3], record[2])
                for record in same_chainage
                if record[sensor + 3] is not None
            ]
            if measurements:
                value = sum(deflection / load * 40 for deflection, load in measurements)
                normalized_deflections.append(value / (len(measurements) * 1000))
            else:
                normalized_deflections.append("NA")
        thickness = 200 if chainage < 49100 else 145
        row = [
            chainage,
            40000,
            0.56,
            *normalized_deflections,
            thickness,
            500,
            *ranges,
            None,
            None,
            None,
        ]
        rows.append((row, same_chainage[0][13], len(same_chainage)))
    return rows


def run(input_file, workbook_path):
    result = get_fwd_data(input_file)
    rows = create_rows(result)

    import xlwings as xw

    workbook_path = Path(workbook_path).resolve()
    with xw.App(visible=False) as app:
        if workbook_path.exists():
            workbook = app.books.open(str(workbook_path))
            LOGGER.info("Opened workbook: %s", workbook_path)
        else:
            workbook = app.books.add()
            workbook.save(str(workbook_path))
            LOGGER.info("Created workbook: %s", workbook_path)
        try:
            if workbook.api.ReadOnly:
                raise PermissionError(
                    f"Workbook is open read-only: {workbook_path}. "
                    "Close the workbook in Excel and remove any stale Excel lock file before retrying."
                )
            if "RHS" in [sheet.name for sheet in workbook.sheets]:
                sheet = workbook.sheets["RHS"]
            else:
                sheet = workbook.sheets.add("RHS", before=workbook.sheets[0])
                LOGGER.info("Created worksheet: RHS")

            write_headers(sheet, app)
            for row_number, (row, asphalt_temperature, measurement_count) in enumerate(rows, start=3):
                sheet.range(f"A{row_number}:W{row_number}").value = row
                sheet.range(f"AH{row_number}").value = asphalt_temperature
                sheet.range(f"AJ{row_number}").value = "Monsoon"
                LOGGER.info(
                    "Wrote output row %d: chainage=%s, grouped measurements=%d",
                    row_number,
                    row[0],
                    measurement_count,
                )
            workbook.save()
            LOGGER.info("Saved %d output rows to %s (sheet RHS)", len(rows), workbook_path)
        finally:
            workbook.close()


def main():
    parser = argparse.ArgumentParser(
        description="Import FWD deflection measurements into the Modulus Calculation workbook."
    )
    parser.add_argument(
        "--input",
        default=SCRIPT_DIR / "MANOHARPUR TO DAUSA RHS.fwd",
        type=Path,
        help="Input FWD file (default: MANOHARPUR TO DAUSA RHS.fwd)",
    )
    parser.add_argument(
        "--workbook",
        default=SCRIPT_DIR / "Modulus_Calculation.xlsx",
        type=Path,
        help="Output Excel workbook (default: Modulus_Calculation.xlsx)",
    )
    args = parser.parse_args()
    LOGGER.info("Starting FWD import; common log: %s", LOG_PATH)
    try:
        run(args.input, args.workbook)
    except Exception:
        LOGGER.exception("FWD import failed")
        raise
    LOGGER.info("FWD import completed successfully")


if __name__ == "__main__":
    main()
