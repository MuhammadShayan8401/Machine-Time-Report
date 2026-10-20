
#!/usr/bin/env python3
"""Build the hourly machine time workbook from Daily Production Report PDFs.

Usage:
    python make_report.py PDF_FOLDER [-o out.xlsx]
        [--threshold 10] [--expect "B650 T" "B1000 T"]

Dependencies:
    pip install pdfplumber openpyxl
"""

import argparse
import glob
import os
import re
import sys
from datetime import timedelta

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule


ROW_RE = re.compile(r"^\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}$")
_SPAN_RE = re.compile(r"(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})")


def dur(s):
    """Convert a duration string containing hours, minutes and seconds to seconds."""
    s = (s or "").strip()

    def get(unit):
        match = re.search(r"(\d+)\s*" + unit, s)
        return int(match.group(1)) if match else 0

    return get("h") * 3600 + get("m") * 60 + get("s")


def norm_name(name):
    """Normalize whitespace in a machine name."""
    return re.sub(r"\s+", " ", re.sub(r"\(", " (", name or "")).strip()


def parse_pdf(path):
    """Extract the machine name, report date and hourly rows from a PDF."""
    machine = None
    date = None
    rows = []

    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""

            if machine is None:
                match = re.search(
                    r"Daily Production Report\s+(.+?)\s*\|", text
                )
                if match:
                    machine = norm_name(match.group(1))

            if date is None:
                match = re.search(r"(\d{1,2} [A-Za-z]{3} \d{4})", text)
                if match:
                    date = match.group(1)

            for table in page.extract_tables():
                for row in table:
                    if (
                        row
                        and row[0]
                        and ROW_RE.match(row[0].strip())
                        and len(row) >= 9
                    ):
                        rows.append(
                            {
                                "time": row[0].strip(),
                                "status": row[1],
                                "shots": row[2],
                                "on": dur(row[6]),
                                "idle": dur(row[7]),
                                "off": dur(row[8]),
                            }
                        )

    return machine, date, rows


def _aggregate_hours(machine_rows):
    """Aggregate source rows by their starting clock hour."""
    slots = [(9 + i) % 24 for i in range(24)]
    aggregate = {hour: [0, 0, 0] for hour in slots}

    for row in machine_rows:
        hour = int(row["time"].split(":")[0]) % 24
        aggregate[hour][0] += row["on"]
        aggregate[hour][1] += row["idle"]
        aggregate[hour][2] += row["off"]

    for hour in slots:
        if sum(aggregate[hour]) == 0:
            aggregate[hour][2] = 3600

    return slots, aggregate


def hourly_table(data, order=None):
    """Return hourly report rows with durations expressed in seconds."""
    output = []

    for machine in (order if order is not None else sorted(data)):
        if machine not in data:
            continue

        slots, aggregate = _aggregate_hours(data[machine])

        for hour in slots:
            on, idle, off = aggregate[hour]
            missing = max(0, 3600 - (on + idle + off))

            output.append(
                {
                    "hour": (
                        f"{hour:02d}:00 - {(hour + 1) % 24:02d}:00"
                    ),
                    "shift": "A" if 9 <= hour < 21 else "B",
                    "machine": machine,
                    "on": on,
                    "idle": idle,
                    "off": off,
                    "missing": missing,
                }
            )

    return output


def build_workbook(data, date, out, threshold=10, expect=()):
    """Build an Excel workbook containing Hourly and Summary worksheets."""
    order = sorted(set(data) | set(expect))
    slots = [(9 + i) % 24 for i in range(24)]

    font_name = "Arial"
    header_font = Font(name=font_name, bold=True, color="FFFFFF")
    header_fill = PatternFill(fill_type="solid", fgColor="1F3864")
    body_font = Font(name=font_name)
    blue_font = Font(name=font_name, color="0000FF")

    thin_side = Side(style="thin", color="BFBFBF")
    border = Border(
        left=thin_side,
        right=thin_side,
        top=thin_side,
        bottom=thin_side,
    )
    duration_format = "[h]:mm:ss"

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Hourly"

    worksheet["A1"] = (
        f"Hourly Machine Time Report - {date or 'Unknown date'} "
        "09:00 to next day 09:00"
    )
    worksheet["A1"].font = Font(name=font_name, bold=True, size=13)

    worksheet["A2"] = "Threshold:"
    worksheet["A2"].font = Font(name=font_name, bold=True)

    worksheet["B2"] = f"=TIME(0,{threshold},0)"
    worksheet["B2"].number_format = duration_format
    worksheet["B2"].font = Font(
        name=font_name, bold=True, color="0000FF"
    )

    headers = [
        "Time Duration",
        "Shift",
        "Machine Name",
        "On Time",
        "Idle Time",
        "Off Time",
        "Missing Time",
        f"Above {threshold} Min Threshold",
    ]

    for column, heading in enumerate(headers, 1):
        cell = worksheet.cell(3, column, heading)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    row_number = 4

    for machine in order:
        if machine not in data:
            continue

        _, aggregate = _aggregate_hours(data[machine])

        for hour in slots:
            on, idle, off = aggregate[hour]

            worksheet.cell(
                row_number, 1,
                f"{hour:02d}:00 - {(hour + 1) % 24:02d}:00",
            )
            worksheet.cell(
                row_number,
                2,
                "A (09:00-21:00)"
                if 9 <= hour < 21
                else "B (21:00-09:00)",
            )
            worksheet.cell(row_number, 3, machine)

            for column, seconds in zip((4, 5, 6), (on, idle, off)):
                cell = worksheet.cell(
                    row_number,
                    column,
                    timedelta(seconds=seconds),
                )
                cell.font = blue_font
                cell.number_format = duration_format

            worksheet.cell(
                row_number,
                7,
                f"=MAX(0,TIME(1,0,0)-(D{row_number}+"
                f"E{row_number}+F{row_number}))",
            )

            worksheet.cell(
                row_number,
                8,
                f'=IF(G{row_number}>$B$2,'
                '"Above threshold","Within threshold")',
            ).alignment = Alignment(horizontal="center")

            for column in range(1, 9):
                cell = worksheet.cell(row_number, column)
                cell.border = border

                if column not in (4, 5, 6):
                    cell.font = body_font

            worksheet.cell(row_number, 7).number_format = duration_format
            row_number += 1

    last_row = row_number - 1

    for column, width in zip(
        "ABCDEFGH", [16, 20, 20, 12, 12, 12, 14, 24]
    ):
        worksheet.column_dimensions[column].width = width

    worksheet.freeze_panes = "A4"

    if last_row >= 4:
        worksheet.auto_filter.ref = f"A3:H{last_row}"

        # Correct PatternFill syntax.
        red_fill = PatternFill(
            fill_type="solid", fgColor="FFC7CE"
        )
        green_fill = PatternFill(
            fill_type="solid", fgColor="C6EFCE"
        )

        worksheet.conditional_formatting.add(
            f"G4:H{last_row}",
            FormulaRule(
                formula=["$G4>$B$2"],
                fill=red_fill,
                font=Font(name=font_name, color="9C0006"),
            ),
        )

        worksheet.conditional_formatting.add(
            f"G4:H{last_row}",
            FormulaRule(
                formula=["$G4<=$B$2"],
                fill=green_fill,
                font=Font(name=font_name, color="006100"),
            ),
        )

    notes_row = max(last_row + 2, 6)
    worksheet.cell(notes_row, 1, "Notes").font = Font(
        name=font_name, bold=True
    )

    notes = [
        "On, Idle and Off Time (blue) are sums of report rows "
        "starting in each clock hour.",
        "An hour with no activity is counted as Off Time = 1:00:00.",
        "Missing Time = max(0, 1 hour - On Time - Idle Time - Off Time).",
    ]

    for index, note in enumerate(notes, 1):
        worksheet.cell(notes_row + index, 1, note).font = body_font

    # Summary worksheet.
    summary = workbook.create_sheet("Summary")

    summary_headers = [
        "Machine Name",
        "On Time",
        "Idle Time",
        "Off Time",
        "Missing Time",
        "Note",
    ]

    for column, heading in enumerate(summary_headers, 1):
        cell = summary.cell(1, column, heading)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    for index, machine in enumerate(order, 2):
        summary.cell(index, 1, machine)

        if machine in data:
            for column, source_column in (
                (2, "D"),
                (3, "E"),
                (4, "F"),
                (5, "G"),
            ):
                cell = summary.cell(
                    index,
                    column,
                    f"=SUMIFS(Hourly!${source_column}$4:"
                    f"${source_column}${last_row},"
                    f"Hourly!$C$4:$C${last_row},$A{index})",
                )
                cell.number_format = duration_format
        else:
            summary.cell(index, 6, "No data in report")

        for column in range(1, 7):
            summary.cell(index, column).font = body_font
            summary.cell(index, column).border = border

    for column, width in zip(
        "ABCDEF", [20, 14, 14, 14, 16, 24]
    ):
        summary.column_dimensions[column].width = width

    workbook.save(out)


def parse_dashboard_text(text):
    """Parse copied dashboard rows, including text without separators.

    Whole-hour rows take precedence over smaller intervals. Shift and
    total rows of four hours or more are ignored.
    """
    duration_pattern = re.compile(
        r"(?:\d+\s*h\s*)?(?:\d+\s*m\s*)?(?:\d+\s*s)?"
    )

    text = (text or "").replace("\r", "")
    text = text.replace("\t", "|").replace("\n", "|")

    spans = list(_SPAN_RE.finditer(text))
    hours = {}
    subhour_rows = []

    for index, match in enumerate(spans):
        segment_end = (
            spans[index + 1].start()
            if index + 1 < len(spans)
            else len(text)
        )
        segment = text[match.end():segment_end]

        shots_match = re.search(r"(\d+)\s*shot\(s\)", segment)
        kwh_match = re.search(r"kWh", segment, re.IGNORECASE)

        if not shots_match or not kwh_match:
            continue

        position = kwh_match.end()
        durations = []

        while len(durations) < 4:
            while position < len(segment) and segment[position] in " |":
                position += 1

            duration_match = duration_pattern.match(segment, position)

            if (
                not duration_match
                or duration_match.end() == position
            ):
                break

            token = duration_match.group().strip()

            # Only accept tokens that contain a duration unit.
            if not re.search(r"[hms]", token):
                break

            durations.append(dur(token))
            position = duration_match.end()

        if len(durations) < 3:
            continue

        h1, m1, h2, m2 = map(int, match.groups())

        start = h1 * 60 + m1
        end = h2 * 60 + m2

        if end <= start:
            end += 24 * 60

        span = end - start

        if span >= 240:
            continue

        row = {
            "time": f"{h1:02d}:{m1:02d} - {h2:02d}:{m2:02d}",
            "status": segment[:shots_match.start()].strip(" |,"),
            "shots": f"{shots_match.group(1)} shot(s)",
            "on": durations[0],
            "idle": durations[1],
            "off": durations[2],
        }

        if start % 60 == 0 and span == 60:
            hours.setdefault(h1 % 24, row)
        elif row not in subhour_rows:
            subhour_rows.append(row)

    return list(hours.values()) + [
        row
        for row in subhour_rows
        if int(row["time"][:2]) % 24 not in hours
    ]


def main():
    """Command-line entry point."""
    parser = argparse.ArgumentParser()
    parser.add_argument("folder")
    parser.add_argument("-o", "--out")
    parser.add_argument(
        "--threshold", type=int, default=10, help="Threshold in minutes"
    )
    parser.add_argument(
        "--expect", nargs="*", default=[],
        help="Machines that should appear",
    )

    args = parser.parse_args()

    data = {}
    report_date = None

    for path in sorted(glob.glob(os.path.join(args.folder, "*.pdf"))):
        machine, date, rows = parse_pdf(path)

        if not machine:
            print("Skipped (machine name not found):", path)
            continue

        report_date = report_date or date
        seen = data.setdefault(machine, [])

        for row in rows:
            if row not in seen:
                seen.append(row)

        print(f"{machine}: {len(rows)} rows from {os.path.basename(path)}")

    if not data:
        sys.exit("No reports found")

    output_path = args.out or (
        f"Hourly_Machine_Time_Report_"
        f"{(report_date or 'report').replace(' ', '')}.xlsx"
    )

    build_workbook(
        data,
        report_date,
        output_path,
        args.threshold,
        args.expect,
    )

    print("Saved", output_path)


if __name__ == "__main__":
    main()
