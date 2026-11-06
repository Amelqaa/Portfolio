import json
import os
import re
import math
from collections import defaultdict, OrderedDict
from pathlib import Path
from datetime import datetime
from functools import reduce
from math import gcd

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

CLINGO_FILE = "asp_output.txt"
RULES_FILE = "main_rules.lp"
DATA_FILE = "data.lp"
OUTPUT_FILE = "schedule.xlsx"

# New shift times in minutes
# AM: 6:00 AM - 2:30 PM
AM_SHIFT_START_MINUTE = 6 * 60
AM_SHIFT_END_MINUTE = 14 * 60 + 30

# PM: 12:00 PM - 8:30 PM
PM_SHIFT_START_MINUTE = 12 * 60
PM_SHIFT_END_MINUTE = 20 * 60 + 30

BREAK_1_DURATION_MINUTES = 15
BREAK_2_DURATION_MINUTES = 30
FULL_SHIFT_DISPLAY_MINUTES = 360

LEVEL_DISPLAY = {
    "t1": "Tech 1",
    "3": "Tech 1",
    "t2": "Tech 2",
    "2": "Tech 2",
    "prev_main": "Prev Main",
    "1": "Prev Main",
    "ld_tech": "Lead Tech",
    "4": "Lead Tech",
}

DEFAULT_DAY_NAMES = {
    0: "Monday",
    1: "Tuesday",
    2: "Wednesday",
    3: "Thursday",
    4: "Friday",
    5: "Saturday",
    6: "Sunday",
}

SHORT_DAY_NAMES = {
    "Monday": "Mon",
    "Tuesday": "Tue",
    "Wednesday": "Wed",
    "Thursday": "Thu",
    "Friday": "Fri",
    "Saturday": "Sat",
    "Sunday": "Sun",
}

DOW_DISPLAY = {
    "mon": "Monday",
    "monday": "Monday",
    "tue": "Tuesday",
    "tues": "Tuesday",
    "tuesday": "Tuesday",
    "wed": "Wednesday",
    "wednesday": "Wednesday",
    "thu": "Thursday",
    "thur": "Thursday",
    "thurs": "Thursday",
    "thursday": "Thursday",
    "fri": "Friday",
    "friday": "Friday",
    "sat": "Saturday",
    "saturday": "Saturday",
    "sun": "Sunday",
    "sunday": "Sunday",
}

# Higher-contrast palette so sections do not blend together
TITLE_FILL = "5B9BD5"
INFO_FILL = "D9EAF7"
HEADER_NAME_FILL = "9DC3E6"
HEADER_META_FILL = "DDEBF7"
JOB_FILL = "9DC3E6"
BREAK_FILL = "FFD966"
TIME_FILL = "F2F2F2"
EMPTY_FILL = "D9D9D9"
WHITE_FILL = "FFFFFF"
SHIFT_FILL = "D9D9D9"
AM_TAB_COLOR = "5B9BD5"
PM_TAB_COLOR = "FFD966"
COMBINED_TAB_COLOR = "70AD47"

THIN_COLOR = "A6A6A6"
MEDIUM_COLOR = "595959"


DEFAULT_BREAKS_BY_SHIFT = {"am": [], "pm": []}


def tech_display_name(tech_id: str) -> str:
    parts = tech_id.strip().split("_")
    output = []
    for part in parts:
        if part.lower() == "jr":
            output.append("Jr")
        else:
            output.append(part.capitalize())
    return " ".join(output)


def parse_atoms(text: str):
    return re.findall(r"([a-zA-Z_][a-zA-Z0-9_]*\([^)]*\))", text)


def clean_comments(text: str) -> str:
    clean_lines = []
    for line in text.splitlines():
        if "%" in line:
            line = line.split("%", 1)[0]
        clean_lines.append(line)
    return "\n".join(clean_lines)


def is_ground_symbol(value: str) -> bool:
    text = str(value or "").strip().strip('"')
    return bool(re.fullmatch(r"[a-z][a-z0-9_]*", text))


def is_ground_date_symbol(value: str) -> bool:
    text = str(value or "").strip().strip('"')
    return bool(re.fullmatch(r"d\d{4}_\d{2}_\d{2}", text))


def minute_to_row(minute_value, timeline_start_row, sheet_start_minute, slot_minutes):
    return timeline_start_row + ((minute_value - sheet_start_minute) // slot_minutes)


def minute_to_row_exclusive(minute_value, timeline_start_row, sheet_start_minute, slot_minutes):
    offset_minutes = minute_value - sheet_start_minute
    return timeline_start_row + ((offset_minutes + slot_minutes - 1) // slot_minutes)


def time_label_full(minute_value):
    h = minute_value // 60
    m = minute_value % 60
    suffix = "AM" if h < 12 else "PM"
    display_h = h % 12
    if display_h == 0:
        display_h = 12
    return f"{display_h}:{m:02d} {suffix}"


def time_label_hour_only(minute_value):
    h = minute_value // 60
    suffix = "AM" if h < 12 else "PM"
    display_h = h % 12
    if display_h == 0:
        display_h = 12
    return f"{display_h}{suffix}"


def compute_slot_minutes(day_jobs, breaks=None):
    relevant_values = []
    for job in day_jobs:
        relevant_values.append(job["duration"])
        relevant_values.append(job["start_minute"] % 60)
        relevant_values.append(job["end_minute"] % 60)

    if breaks:
        for br in breaks:
            relevant_values.append(br["start"] % 60)
            relevant_values.append(br["end"] % 60)
            relevant_values.append(br["end"] - br["start"])

    relevant_values = [value for value in relevant_values if value > 0]

    if not relevant_values:
        return 15

    for candidate in [30, 20, 15, 12, 10, 6, 5, 4, 3, 2, 1]:
        if all(value % candidate == 0 for value in relevant_values):
            return candidate

    return 1


def format_total_minutes(total_minutes):
    hours = total_minutes // 60
    minutes = total_minutes % 60
    if minutes == 0:
        return f"{hours}h"
    return f"{hours}h {minutes}m"


def format_total_out_of_shift(total_minutes):
    return f"{total_minutes} min ({format_total_minutes(total_minutes)} / {format_total_minutes(FULL_SHIFT_DISPLAY_MINUTES)})"


def choose_am_pm(label):
    while True:
        print(f"{label} - choose:")
        print("1. AM")
        print("2. PM")
        choice = input("Enter 1 or 2: ").strip()
        if choice == "1":
            return "am"
        if choice == "2":
            return "pm"
        print("Invalid choice. Try again.\n")


def convert_time_ampm_to_minute(time_value, am_pm):
    time_value = time_value.strip()
    am_pm = am_pm.strip().lower()

    if am_pm not in ["am", "pm"]:
        raise ValueError("Type must be AM or PM.")

    if ":" in time_value:
        parts = time_value.split(":")
        if len(parts) != 2:
            raise ValueError("Time must look like 9 or 9:15 or 2:30.")
        hour_part, minute_part = parts[0].strip(), parts[1].strip()
        if not hour_part.isdigit() or not minute_part.isdigit():
            raise ValueError("Time must contain numbers only.")
    else:
        hour_part = time_value
        minute_part = "0"
        if not hour_part.isdigit():
            raise ValueError("Time must contain numbers only.")

    hour_value = int(hour_part)
    minute_value = int(minute_part)

    if hour_value < 1 or hour_value > 12:
        raise ValueError("Hour must be between 1 and 12.")
    if minute_value < 0 or minute_value > 59:
        raise ValueError("Minutes must be between 00 and 59.")

    if am_pm == "am":
        base_hour = 0 if hour_value == 12 else hour_value
    else:
        base_hour = 12 if hour_value == 12 else hour_value + 12

    return base_hour * 60 + minute_value


def validate_break_within_shift(start_minute, duration, shift_start, shift_end, label):
    end_minute = start_minute + duration
    if start_minute < shift_start or end_minute > shift_end:
        raise ValueError(
            f"{label} is outside the shift time. "
            f"It must stay within {time_label_full(shift_start)} - {time_label_full(shift_end)}."
        )


def prompt_validated_am_break(label, duration, shift_start, shift_end):
    while True:
        try:
            time_value = input(
                f"{label} start time within {time_label_full(shift_start)} - {time_label_full(shift_end)} "
                f"(example 9 or 9:15): "
            ).strip()
            am_pm = choose_am_pm(label)
            start_minute = convert_time_ampm_to_minute(time_value, am_pm)
            validate_break_within_shift(start_minute, duration, shift_start, shift_end, label)
            return {"label": label, "start": start_minute, "end": start_minute + duration}
        except ValueError as e:
            print(f"{e}\nPlease try again.\n")


def prompt_validated_pm_break(label, duration, shift_start, shift_end):
    while True:
        try:
            time_value = input(
                f"{label} start time within {time_label_full(shift_start)} - {time_label_full(shift_end)} "
                f"(PM only, example 2 or 2:15): "
            ).strip()
            start_minute = convert_time_ampm_to_minute(time_value, "pm")
            validate_break_within_shift(start_minute, duration, shift_start, shift_end, label)
            return {"label": label, "start": start_minute, "end": start_minute + duration}
        except ValueError as e:
            print(f"{e}\nPlease try again.\n")


def prompt_two_breaks_for_am_shift():
    print("\nEnter 2 break start times for AM shift.")
    print(f"AM shift window: {time_label_full(AM_SHIFT_START_MINUTE)} - {time_label_full(AM_SHIFT_END_MINUTE)}.")
    print("AM Break 1 is 15 minutes. AM Break 2 Lunch is 30 minutes.")


    break_1 = prompt_validated_am_break(
        "AM Break 1", BREAK_1_DURATION_MINUTES, AM_SHIFT_START_MINUTE, AM_SHIFT_END_MINUTE
    )
    break_2 = prompt_validated_am_break(
        "AM Break 2 Lunch", BREAK_2_DURATION_MINUTES, AM_SHIFT_START_MINUTE, AM_SHIFT_END_MINUTE
    )
    breaks = [break_1, break_2]
    breaks.sort(key=lambda x: x["start"])
    if breaks[0]["end"] > breaks[1]["start"]:
        raise ValueError("AM breaks cannot overlap. Please rerun and enter non-overlapping break times.")
    return breaks


def prompt_two_breaks_for_pm_shift():
    print("\nEnter 2 break start times for PM shift.")
    print(f"PM shift window: {time_label_full(PM_SHIFT_START_MINUTE)} - {time_label_full(PM_SHIFT_END_MINUTE)}.")
    print("Entered times are treated as PM.")
    print("PM Break 1 is 15 minutes. PM Break 2 Lunch is 30 minutes.")

    break_1 = prompt_validated_pm_break(
        "PM Break 1", BREAK_1_DURATION_MINUTES, PM_SHIFT_START_MINUTE, PM_SHIFT_END_MINUTE
    )
    break_2 = prompt_validated_pm_break(
        "PM Break 2 Lunch", BREAK_2_DURATION_MINUTES, PM_SHIFT_START_MINUTE, PM_SHIFT_END_MINUTE
    )
    breaks = [break_1, break_2]
    breaks.sort(key=lambda x: x["start"])
    if breaks[0]["end"] > breaks[1]["start"]:
        raise ValueError("PM breaks cannot overlap. Please rerun and enter non-overlapping break times.")
    return breaks


def prompt_break_times_by_shift():
    while True:
        try:
            return {
                "am": prompt_two_breaks_for_am_shift(),
                "pm": prompt_two_breaks_for_pm_shift(),
            }
        except ValueError as e:
            print(f"{e}\nPlease enter the break times again.\n")


def format_break_times(breaks):
    return " | ".join(
        f"{item['label']}: {time_label_full(item['start'])} - {time_label_full(item['end'])}"
        for item in breaks
    )


def parse_clingo(path: str):
    with open(path, "r", encoding="utf-8") as f:
        text = f.read()

    atoms = parse_atoms(text)

    data = {
        "assignments": [],
        "seg_duration": {},
        "due_day": {},
        "job_end_day": {},
        "job_start_minutes": {},
        "job_start_hour": {},
    }

    for atom in atoms:
        if atom.startswith("assign_seg("):
            m = re.match(r"assign_seg\(([^,]+),(\d+),([^,]+),(\d+)\)", atom)
            if m:
                job, seg, tech, day = m.groups()
                data["assignments"].append(
                    {
                        "job": job.strip(),
                        "seg": int(seg),
                        "tech": tech.strip(),
                        "day": int(day),
                    }
                )

        elif atom.startswith("seg_duration("):
            m = re.match(r"seg_duration\(([^,]+),(\d+),(\d+)\)", atom)
            if m:
                job, seg, dur = m.groups()
                data["seg_duration"][(job.strip(), int(seg))] = int(dur)

        elif atom.startswith("due_day("):
            m = re.match(r"due_day\(([^,]+),(\d+)\)", atom)
            if m:
                job, day = m.groups()
                data["due_day"][job.strip()] = int(day)

        elif atom.startswith("job_end_day("):
            m = re.match(r"job_end_day\(([^,]+),(\d+)\)", atom)
            if m:
                job, day = m.groups()
                data["job_end_day"][job.strip()] = int(day)

        elif atom.startswith("seg_start_minute("):
            m = re.match(r"seg_start_minute\(([^,]+),(\d+),(\d+)\)", atom)
            if m:
                job, seg, minute = m.groups()
                data["job_start_minutes"][(job.strip(), int(seg))] = int(minute)

        elif atom.startswith("seg_start_hour("):
            m = re.match(r"seg_start_hour\(([^,]+),(\d+),(\d+)\)", atom)
            if m:
                job, seg, hour = m.groups()
                data["job_start_hour"][(job.strip(), int(seg))] = int(hour) * 60

    if not data["assignments"]:
        raise ValueError(f"No assign_seg/4 atoms found in {path}")

    return data


def parse_clingo_text(text: str):
    atoms = parse_atoms(text)

    data = {
        "assignments": [],
        "seg_duration": {},
        "due_day": {},
        "job_end_day": {},
        "job_start_minutes": {},
        "job_start_hour": {},
    }

    for atom in atoms:
        if atom.startswith("assign_seg("):
            m = re.match(r"assign_seg\(([^,]+),(\d+),([^,]+),(\d+)\)", atom)
            if m:
                job, seg, tech, day = m.groups()
                data["assignments"].append(
                    {
                        "job": job.strip(),
                        "seg": int(seg),
                        "tech": tech.strip(),
                        "day": int(day),
                    }
                )
        elif atom.startswith("seg_duration("):
            m = re.match(r"seg_duration\(([^,]+),(\d+),(\d+)\)", atom)
            if m:
                job, seg, dur = m.groups()
                data["seg_duration"][(job.strip(), int(seg))] = int(dur)
        elif atom.startswith("due_day("):
            m = re.match(r"due_day\(([^,]+),(\d+)\)", atom)
            if m:
                job, day = m.groups()
                data["due_day"][job.strip()] = int(day)
        elif atom.startswith("job_end_day("):
            m = re.match(r"job_end_day\(([^,]+),(\d+)\)", atom)
            if m:
                job, day = m.groups()
                data["job_end_day"][job.strip()] = int(day)
        elif atom.startswith("seg_start_minute("):
            m = re.match(r"seg_start_minute\(([^,]+),(\d+),(\d+)\)", atom)
            if m:
                job, seg, minute = m.groups()
                data["job_start_minutes"][(job.strip(), int(seg))] = int(minute)
        elif atom.startswith("seg_start_hour("):
            m = re.match(r"seg_start_hour\(([^,]+),(\d+),(\d+)\)", atom)
            if m:
                job, seg, hour = m.groups()
                data["job_start_hour"][(job.strip(), int(seg))] = int(hour) * 60

    if not data["assignments"]:
        raise ValueError("No assign_seg/4 atoms found in the best optimisation result.")

    return data


def parse_lp_file(path: str):
    data = {
        "tech_level": {},
        "tech_shift": {},
        "techs": OrderedDict(),
        "day_of_week": {},
        "date_day": {},
        "need_time_minutes": {},
        "need_time_date": {},
        "job_labels": {},
    }

    if not os.path.exists(path):
        return data

    with open(path, "r", encoding="utf-8") as f:
        text = f.read()

    clean_text = clean_comments(text)
    atoms = parse_atoms(clean_text)

    for atom in atoms:
        if atom.startswith("tech("):
            m = re.match(r"tech\(([^,]+),([^)]+)\)", atom)
            if m:
                tech, shift = m.groups()
                tech = tech.strip()
                shift = shift.strip().strip('"').lower()
                if is_ground_symbol(tech) and is_ground_symbol(shift):
                    data["techs"][tech] = shift
                    data["tech_shift"][tech] = shift

        elif atom.startswith("tech_level("):
            m = re.match(r"tech_level\(([^,]+),([^)]+)\)", atom)
            if m:
                tech, level = m.groups()
                tech = tech.strip()
                level = level.strip().strip('"').lower()
                if is_ground_symbol(tech) and is_ground_symbol(level):
                    data["tech_level"][tech] = level

        elif atom.startswith("day_of_week("):
            m = re.match(r"day_of_week\((\d+),([^)]+)\)", atom)
            if m:
                day, dow = m.groups()
                dow = dow.strip().strip('"').lower()
                if is_ground_symbol(dow):
                    data["day_of_week"][int(day)] = dow

        elif atom.startswith("date_day("):
            m = re.match(r"date_day\((\d+),([^)]+)\)", atom)
            if m:
                day, date_value = m.groups()
                date_value = date_value.strip().strip('"')
                if is_ground_date_symbol(date_value):
                    data["date_day"][int(day)] = date_value

        elif atom.startswith("need_time("):
            m = re.match(r"need_time\(([^,]+),([^,]+),([^)]+)\)", atom)
            if m:
                job, date_value, minutes = m.groups()
                job = job.strip()
                date_value = date_value.strip().strip('"')
                minutes = minutes.strip().strip('"')
                if is_ground_symbol(job) and is_ground_date_symbol(date_value):
                    data["need_time_date"][job] = date_value
                if is_ground_symbol(job) and minutes.isdigit():
                    data["need_time_minutes"][job] = int(minutes)

        elif atom.startswith("job_label("):
            m = re.match(r'job_label\(([^,]+),\s*(\".*\")\)', atom)
            if m:
                job, label_literal = m.groups()
                job = job.strip()
                if is_ground_symbol(job):
                    try:
                        data["job_labels"][job] = json.loads(label_literal)
                    except json.JSONDecodeError:
                        data["job_labels"][job] = label_literal.strip().strip('"')

    return data


def merge_metadata(main_rules, data_file):
    merged = {
        "tech_level": {},
        "tech_shift": {},
        "techs": OrderedDict(),
        "day_of_week": {},
        "date_day": {},
        "need_time_minutes": {},
        "need_time_date": {},
        "job_labels": {},
    }

    merged["techs"].update(main_rules["techs"])
    merged["techs"].update(data_file["techs"])

    merged["tech_level"].update(main_rules["tech_level"])
    merged["tech_level"].update(data_file["tech_level"])

    merged["tech_shift"].update(main_rules["tech_shift"])
    merged["tech_shift"].update(data_file["tech_shift"])

    merged["day_of_week"].update(main_rules["day_of_week"])
    merged["day_of_week"].update(data_file["day_of_week"])

    merged["date_day"].update(main_rules["date_day"])
    merged["date_day"].update(data_file["date_day"])

    merged["need_time_minutes"].update(main_rules["need_time_minutes"])
    merged["need_time_minutes"].update(data_file["need_time_minutes"])

    merged["need_time_date"].update(main_rules["need_time_date"])
    merged["need_time_date"].update(data_file["need_time_date"])

    merged["job_labels"].update(main_rules["job_labels"])
    merged["job_labels"].update(data_file["job_labels"])

    return merged


def build_tech_metadata(meta):
    if not meta["techs"]:
        raise ValueError("No tech(...) facts found in data.lp or main_rules.lp")

    tech_meta = OrderedDict()

    for tech in sorted(meta["techs"].keys()):
        shift = meta["techs"][tech]
        raw_level = meta["tech_level"].get(tech, "")
        tech_meta[tech] = {
            "shift": shift,
            "level_raw": raw_level,
            "level_display": LEVEL_DISPLAY.get(raw_level, raw_level if raw_level else "Unknown"),
            "name": tech_display_name(tech),
        }

    return tech_meta


def get_days(clingo):
    return sorted({item["day"] for item in clingo["assignments"]})


def get_day_name(day, meta):
    day_name = DEFAULT_DAY_NAMES.get(day, f"Day {day}")
    if day in meta["day_of_week"]:
        raw = meta["day_of_week"][day].lower()
        day_name = DOW_DISPLAY.get(raw, day_name)
    return day_name


def get_date_for_day(day, meta):
    value = str(meta["date_day"].get(day, str(day))).strip()
    m = re.search(r"\d{4}-\d{2}-\d{2}", value)
    if m:
        return m.group()
    return value


def get_job_display(job, meta):
    label = str(meta.get("job_labels", {}).get(job, "")).strip()
    if not label:
        return job
    return f"{label} ({job})"


def get_original_item_sort_key(clingo, item):
    key = (item["job"], item["seg"])
    if key in clingo["job_start_minutes"]:
        start_minute = clingo["job_start_minutes"][key]
    elif key in clingo["job_start_hour"]:
        start_minute = clingo["job_start_hour"][key]
    else:
        start_minute = 10**9
    return (start_minute, item["seg"], item["job"])


def group_assignments(clingo):
    grouped = defaultdict(list)
    for item in clingo["assignments"]:
        grouped[(item["day"], item["tech"])].append(item)
    for key in grouped:
        grouped[key] = sorted(grouped[key], key=lambda x: get_original_item_sort_key(clingo, x))
    return grouped


def shift_start_minute_for_tech(shift_name):
    return PM_SHIFT_START_MINUTE if shift_name == "pm" else AM_SHIFT_START_MINUTE


def shift_end_minute_for_tech(shift_name):
    return PM_SHIFT_END_MINUTE if shift_name == "pm" else AM_SHIFT_END_MINUTE


def push_time_out_of_breaks(start_minute, breaks):
    current = start_minute
    while True:
        moved = False
        for br in breaks:
            if br["start"] <= current < br["end"]:
                current = br["end"]
                moved = True
        if not moved:
            return current


def split_job_around_breaks(start_minute, duration, breaks):
    segments = []
    current = push_time_out_of_breaks(start_minute, breaks)
    remaining = duration

    while remaining > 0:
        next_break = None
        for br in breaks:
            if br["start"] >= current:
                next_break = br
                break

        if next_break is None:
            segments.append({"start_minute": current, "end_minute": current + remaining, "duration": remaining})
            break

        work_window = next_break["start"] - current
        if work_window <= 0:
            current = push_time_out_of_breaks(next_break["end"], breaks)
            continue

        worked = min(remaining, work_window)
        segments.append({"start_minute": current, "end_minute": current + worked, "duration": worked})
        remaining -= worked
        current += worked

        if remaining > 0 and current == next_break["start"]:
            current = push_time_out_of_breaks(next_break["end"], breaks)

    return segments


def build_day_lane_schedule(clingo, techs, days, tech_meta, breaks_by_shift):
    grouped = group_assignments(clingo)
    result = defaultdict(lambda: defaultdict(list))

    for day in days:
        for tech in techs:
            items = grouped.get((day, tech), [])
            shift_name = tech_meta[tech]["shift"]
            breaks = breaks_by_shift.get(shift_name, [])
            cursor = shift_start_minute_for_tech(shift_name)
            cursor = push_time_out_of_breaks(cursor, breaks)

            for entry in items:
                job = entry["job"]
                seg = entry["seg"]
                duration = clingo["seg_duration"].get((job, seg), 60)

                pieces = split_job_around_breaks(cursor, duration, breaks)

                for part_index, piece in enumerate(pieces, start=1):
                    result[day][shift_name].append(
                        {
                            "tech": tech,
                            "job": job,
                            "seg": seg,
                            "part": part_index,
                            "parts_total": len(pieces),
                            "start_minute": piece["start_minute"],
                            "end_minute": piece["end_minute"],
                            "duration": piece["duration"],
                            "shift": shift_name,
                            "is_break": False,
                        }
                    )

                if pieces:
                    cursor = pieces[-1]["end_minute"]
                    cursor = push_time_out_of_breaks(cursor, breaks)

            for br in breaks:
                result[day][shift_name].append(
                    {
                        "tech": tech,
                        "job": br["label"],
                        "seg": 0,
                        "part": 1,
                        "parts_total": 1,
                        "start_minute": br["start"],
                        "end_minute": br["end"],
                        "duration": br["end"] - br["start"],
                        "shift": shift_name,
                        "is_break": True,
                    }
                )

            result[day][shift_name].sort(key=lambda x: (x["tech"], x["start_minute"], x["is_break"]))

    return result


def build_styles(shift_name):
    thin = Side(style="thin", color=THIN_COLOR)
    medium = Side(style="medium", color=MEDIUM_COLOR)

    if shift_name == "combined":
        tab_color = COMBINED_TAB_COLOR
    elif shift_name == "pm":
        tab_color = PM_TAB_COLOR
    else:
        tab_color = AM_TAB_COLOR

    return {
        "title_fill": PatternFill("solid", fgColor=TITLE_FILL),
        "info_fill": PatternFill("solid", fgColor=INFO_FILL),
        "header_name_fill": PatternFill("solid", fgColor=HEADER_NAME_FILL),
        "header_meta_fill": PatternFill("solid", fgColor=HEADER_META_FILL),
        "time_fill": PatternFill("solid", fgColor=TIME_FILL),
        "shift_fill": PatternFill("solid", fgColor=SHIFT_FILL),
        "job_fill": PatternFill("solid", fgColor=JOB_FILL),
        "break_fill": PatternFill("solid", fgColor=BREAK_FILL),
        "empty_fill": PatternFill("solid", fgColor=EMPTY_FILL),
        "tab_color": tab_color,

        "thin_border": Border(left=thin, right=thin, top=thin, bottom=thin),
        "medium_border": Border(left=medium, right=medium, top=medium, bottom=medium),

        "title_font": Font(bold=True, size=9, color="1F1F1F"),
        "header_font": Font(bold=True, size=7, color="1F1F1F"),
        "subheader_font": Font(bold=True, size=6, color="404040"),
        "job_font": Font(bold=True, size=6, color="1F1F1F"),
        "break_font": Font(bold=True, size=6, color="7F6000"),
        "normal_font": Font(size=6, color="404040"),

        "center": Alignment(horizontal="center", vertical="center", wrap_text=True),
        "left": Alignment(horizontal="left", vertical="center", wrap_text=True),
        "top_left": Alignment(horizontal="left", vertical="top", wrap_text=True, shrink_to_fit=True),
    }


def style_range(ws, start_row, end_row, start_col, end_col, fill=None, border=None, alignment=None, font=None):
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(r, c)
            if fill:
                cell.fill = fill
            if border:
                cell.border = border
            if alignment:
                try:
                    cell.alignment = alignment
                except AttributeError:
                    pass
            if font:
                try:
                    cell.font = font
                except AttributeError:
                    pass


def ensure_row_min_height(ws, row, min_height):
    current_height = ws.row_dimensions[row].height
    if current_height is None or current_height < min_height:
        ws.row_dimensions[row].height = min_height


def merged_width(ws, start_col, end_col):
    total = 0
    for col in range(start_col, end_col + 1):
        letter = get_column_letter(col)
        width = ws.column_dimensions[letter].width
        total += width if width is not None else 8.43
    return total


def wrapped_line_count(text, available_width):
    if available_width <= 1:
        available_width = 1

    estimated_chars_per_line = max(1, int(available_width * 0.9))
    total_lines = 0
    for raw_line in str(text).splitlines() or [""]:
        line = raw_line if raw_line else " "
        total_lines += max(1, math.ceil(len(line) / estimated_chars_per_line))
    return total_lines


def min_height_for_text(ws, text, start_col, end_col=None, is_break=False):
    if end_col is None:
        end_col = start_col

    available_width = merged_width(ws, start_col, end_col)
    line_count = wrapped_line_count(text, available_width)
    per_line = 7 if is_break else 8
    base = 4 if is_break else 5
    return base + (line_count * per_line)


def remove_overlapping_merges(ws, col, start_row, end_row):
    overlaps = []
    for merged_range in list(ws.merged_cells.ranges):
        same_col = merged_range.min_col <= col <= merged_range.max_col
        rows_overlap = not (end_row < merged_range.min_row or start_row > merged_range.max_row)
        if same_col and rows_overlap:
            overlaps.append(str(merged_range))
    for rng in overlaps:
        ws.unmerge_cells(rng)


def block_border(styles, row, start_row, end_row):
    medium = styles["medium_border"]
    none = Side(style=None)
    return Border(
        left=medium.left,
        right=medium.right,
        top=medium.top if row == start_row else none,
        bottom=medium.bottom if row == end_row else none,
    )


def side_border_only(styles):
    medium = styles["medium_border"]
    none = Side(style=None)
    return Border(left=medium.left, right=medium.right, top=none, bottom=none)


def draw_block(ws, col, start_row, end_row, text, styles, is_break=False):
    remove_overlapping_merges(ws, col, start_row, end_row)

    fill = styles["break_fill"] if is_break else styles["job_fill"]
    font = styles["break_font"] if is_break else styles["job_font"]

    for r in range(start_row, end_row + 1):
        cell = ws.cell(r, col)
        cell.fill = fill
        cell.border = block_border(styles, r, start_row, end_row)
        cell.value = None
        cell.alignment = styles["center"] if is_break else styles["top_left"]

    anchor = ws.cell(start_row, col)
    anchor.value = text
    anchor.font = font
    anchor.alignment = styles["center"] if is_break else styles["top_left"]
    ensure_row_min_height(ws, start_row, min_height_for_text(ws, text, col, is_break=is_break))


def draw_empty_merged_block(ws, col, start_row, end_row, styles):
    if end_row < start_row:
        return

    remove_overlapping_merges(ws, col, start_row, end_row)

    for r in range(start_row, end_row + 1):
        cell = ws.cell(r, col)
        cell.fill = styles["shift_fill"]
        cell.border = side_border_only(styles)
        cell.value = None
        cell.alignment = styles["top_left"]


def is_weekend(day_name):
    return day_name in {"Saturday", "Sunday"}


def build_lane_segments(start_minute, end_minute, blocks):
    segments = []
    current = start_minute
    blocks = sorted(blocks, key=lambda x: x["start_minute"])

    for block in blocks:
        b_start = max(block["start_minute"], start_minute)
        b_end = min(block["end_minute"], end_minute)

        if b_end <= b_start:
            continue

        if b_start > current:
            segments.append({
                "type": "empty",
                "start_minute": current,
                "end_minute": b_start,
            })

        segments.append({
            "type": "break" if block["is_break"] else "job",
            **block,
            "start_minute": b_start,
            "end_minute": b_end,
        })

        current = b_end

    if current < end_minute:
        segments.append({
            "type": "empty",
            "start_minute": current,
            "end_minute": end_minute,
        })

    return segments


def normalize_tech_blocks_for_display(tech, shift_name, tech_blocks, shift_start, breaks):
    normalized = []
    cursor = push_time_out_of_breaks(shift_start, breaks)
    job_blocks = sorted(
        [block for block in tech_blocks if not block["is_break"]],
        key=lambda block: (block["start_minute"], block["seg"], block.get("part", 1), block["job"]),
    )

    for block in job_blocks:
        pieces = split_job_around_breaks(cursor, block["duration"], breaks)
        generated_parts_total = len(pieces)

        for generated_part_index, piece in enumerate(pieces, start=1):
            normalized.append(
                {
                    **block,
                    "start_minute": piece["start_minute"],
                    "end_minute": piece["end_minute"],
                    "duration": piece["duration"],
                    "part": generated_part_index if generated_parts_total > 1 else block.get("part", 1),
                    "parts_total": (
                        generated_parts_total if generated_parts_total > 1 else block.get("parts_total", 1)
                    ),
                }
            )

        if pieces:
            cursor = push_time_out_of_breaks(pieces[-1]["end_minute"], breaks)

    for br in breaks:
        normalized.append(
            {
                "tech": tech,
                "job": br["label"],
                "seg": 0,
                "part": 1,
                "parts_total": 1,
                "start_minute": br["start"],
                "end_minute": br["end"],
                "duration": br["end"] - br["start"],
                "shift": shift_name,
                "is_break": True,
            }
        )

    normalized.sort(key=lambda block: (block["start_minute"], block["is_break"], block["job"], block["seg"]))
    return normalized


def draw_shift_sheet(ws, day, shift_name, meta, tech_meta, jobs_for_shift, styles, breaks, breaks_by_shift=None):
    if shift_name == "combined":
        techs = list(tech_meta.keys())
    else:
        techs = [tech for tech in tech_meta.keys() if tech_meta[tech]["shift"] == shift_name]

    if not techs:
        return

    if shift_name == "combined":
        start_minute = AM_SHIFT_START_MINUTE
        end_minute = PM_SHIFT_END_MINUTE
    else:
        start_minute = shift_start_minute_for_tech(shift_name)
        end_minute = shift_end_minute_for_tech(shift_name)
    slot_minutes = compute_slot_minutes(jobs_for_shift, breaks)
    total_slots = max(1, minute_to_row_exclusive(end_minute, 0, start_minute, slot_minutes))

    title_row = 1
    info_row = 2
    breaks_row = 3
    header_row_1 = 4
    header_row_2 = 5
    header_row_3 = 6
    timeline_start_row = 7

    last_col = 1 + len(techs)

    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = styles["tab_color"]

    ws.column_dimensions["A"].width = 3.6
    tech_col_width = 18 if shift_name == "combined" else 24
    for idx, tech in enumerate(techs, start=2):
        ws.column_dimensions[get_column_letter(idx)].width = tech_col_width

    ws.row_dimensions[title_row].height = 9
    ws.row_dimensions[info_row].height = 9
    ws.row_dimensions[breaks_row].height = 10
    ws.row_dimensions[header_row_1].height = 9
    ws.row_dimensions[header_row_2].height = 8
    ws.row_dimensions[header_row_3].height = 8

    for r in range(timeline_start_row, timeline_start_row + total_slots):
        ws.row_dimensions[r].height = 4.8

    day_name = get_day_name(day, meta)
    sheet_date = get_date_for_day(day, meta)
    try:
        dt = datetime.strptime(sheet_date, "%Y-%m-%d")
        display_date = f"{dt.month}/{dt.day}"
    except ValueError:
        display_date = sheet_date

    title_label = "AM+PM" if shift_name == "combined" else f"{shift_name.upper()} SHIFT"
    title_text = f"{day_name} ({display_date}) - {title_label}"

    ws.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=last_col)
    c = ws.cell(title_row, 1)
    c.value = title_text
    c.fill = styles["title_fill"]
    c.font = styles["title_font"]
    c.alignment = styles["center"]
    ensure_row_min_height(ws, title_row, min_height_for_text(ws, title_text, 1, last_col))
    style_range(ws, title_row, title_row, 1, last_col, border=styles["medium_border"])

    total_jobs = len([job for job in jobs_for_shift if not job["is_break"]])
    ws.merge_cells(start_row=info_row, start_column=1, end_row=info_row, end_column=last_col)
    c = ws.cell(info_row, 1)
    c.value = (
        f"Techs: {len(techs)}    "
        f"Jobs: {total_jobs}    "
        f"{time_label_full(start_minute)} - {time_label_full(end_minute)}"
    )
    c.alignment = styles["left"]
    c.font = styles["normal_font"]
    c.fill = styles["info_fill"]
    ensure_row_min_height(ws, info_row, min_height_for_text(ws, c.value, 1, last_col))
    style_range(ws, info_row, info_row, 1, last_col, border=styles["thin_border"])

    ws.merge_cells(start_row=breaks_row, start_column=1, end_row=breaks_row, end_column=last_col)
    c = ws.cell(breaks_row, 1)
    c.value = f"Breaks: {format_break_times(breaks)}"
    c.alignment = styles["left"]
    c.font = styles["normal_font"]
    c.fill = styles["break_fill"]
    ensure_row_min_height(ws, breaks_row, min_height_for_text(ws, c.value, 1, last_col, is_break=True))
    style_range(ws, breaks_row, breaks_row, 1, last_col, border=styles["thin_border"])

    c = ws.cell(header_row_1, 1)
    c.value = "Time"
    c.fill = styles["header_name_fill"]
    c.font = styles["header_font"]
    c.alignment = styles["center"]
    c.border = styles["medium_border"]
    ensure_row_min_height(ws, header_row_1, min_height_for_text(ws, c.value, 1))

    c = ws.cell(header_row_2, 1)
    c.value = ""
    c.fill = styles["header_meta_fill"]
    c.border = styles["medium_border"]

    c = ws.cell(header_row_3, 1)
    c.value = "Total Assigned"
    c.fill = styles["header_meta_fill"]
    c.font = styles["header_font"]
    c.alignment = styles["center"]
    c.border = styles["medium_border"]
    ensure_row_min_height(ws, header_row_3, min_height_for_text(ws, c.value, 1))

    for idx, tech in enumerate(techs, start=2):
        meta_row = tech_meta[tech]
        total_minutes = sum(
            job["duration"] for job in jobs_for_shift if job["tech"] == tech and not job["is_break"]
        )

        top = ws.cell(header_row_1, idx)
        top.value = meta_row["name"].upper()
        top.fill = styles["header_name_fill"]
        top.font = styles["header_font"]
        top.alignment = styles["center"]
        top.border = styles["medium_border"]
        ensure_row_min_height(ws, header_row_1, min_height_for_text(ws, top.value, idx))

        middle = ws.cell(header_row_2, idx)
        middle.value = f"{meta_row['level_display'].upper()} | {meta_row['shift'].upper()}"
        middle.fill = styles["header_meta_fill"]
        middle.font = styles["subheader_font"]
        middle.alignment = styles["center"]
        middle.border = styles["medium_border"]
        ensure_row_min_height(ws, header_row_2, min_height_for_text(ws, middle.value, idx))

        bottom = ws.cell(header_row_3, idx)
        bottom.value = format_total_out_of_shift(total_minutes)
        bottom.fill = styles["header_meta_fill"]
        bottom.font = styles["subheader_font"]
        bottom.alignment = styles["center"]
        bottom.border = styles["medium_border"]
        ensure_row_min_height(ws, header_row_3, min_height_for_text(ws, bottom.value, idx))

    current_minute = start_minute
    for slot in range(total_slots):
        row = timeline_start_row + slot

        t = ws.cell(row, 1)
        t.value = time_label_hour_only(current_minute) if current_minute % 60 == 0 else ""
        t.fill = styles["time_fill"]
        t.font = styles["header_font"]
        t.alignment = styles["center"]
        t.border = Border()
        if t.value:
            ensure_row_min_height(ws, row, min_height_for_text(ws, t.value, 1))

        for tech_idx, _tech in enumerate(techs, start=2):
            cell = ws.cell(row, tech_idx)
            cell.fill = styles["shift_fill"]
            cell.border = side_border_only(styles)
            cell.alignment = styles["top_left"]

        current_minute += slot_minutes

    tech_col_map = {tech: idx for idx, tech in enumerate(techs, start=2)}

    for tech in techs:
        col = tech_col_map[tech]
        tech_blocks = [b for b in jobs_for_shift if b["tech"] == tech]
        tech_shift_name = tech_meta[tech]["shift"] if shift_name == "combined" else shift_name
        tech_start_minute = shift_start_minute_for_tech(tech_shift_name)
        tech_breaks = (
            breaks_by_shift.get(tech_shift_name, [])
            if shift_name == "combined" and breaks_by_shift
            else breaks
        )
        normalized_blocks = normalize_tech_blocks_for_display(
            tech,
            tech_shift_name,
            tech_blocks,
            tech_start_minute,
            tech_breaks,
        )
        lane_segments = build_lane_segments(start_minute, end_minute, normalized_blocks)

        for seg in lane_segments:
            seg_start = seg["start_minute"]
            seg_end = seg["end_minute"]
            if seg_end <= seg_start:
                continue

            start_row = minute_to_row(seg_start, timeline_start_row, start_minute, slot_minutes)
            end_row = minute_to_row_exclusive(seg_end, timeline_start_row, start_minute, slot_minutes) - 1

            if seg["type"] == "empty":
                draw_empty_merged_block(ws, col, start_row, end_row, styles)
            elif seg["type"] == "break":
                text = (
                    f"{seg['job']}\n"
                    f"{time_label_full(seg['start_minute'])} - {time_label_full(seg['end_minute'])}"
                )
                draw_block(ws, col, start_row, end_row, text, styles, is_break=True)
            else:
                part_label = f" Part {seg['part']}/{seg['parts_total']}" if seg["parts_total"] > 1 else ""
                job_display = get_job_display(seg["job"], meta)
                text = (
                    f"{job_display}\n"
                    f"Seg {seg['seg']}{part_label}\n"
                    f"{time_label_full(seg['start_minute'])} - {time_label_full(seg['end_minute'])}\n"
                    f"{seg['duration']} min"
                )
                draw_block(ws, col, start_row, end_row, text, styles, is_break=False)

    minutes_per_hour = 60
    slots_per_hour = max(1, minutes_per_hour // slot_minutes)
    current_minute = start_minute
    hour_row_start = timeline_start_row
    while current_minute < end_minute:
        hour_row_end = min(hour_row_start + slots_per_hour - 1, timeline_start_row + total_slots - 1)

        if hour_row_end > hour_row_start:
            ws.merge_cells(start_row=hour_row_start, start_column=1, end_row=hour_row_end, end_column=1)

        time_anchor = ws.cell(hour_row_start, 1)
        time_anchor.value = time_label_hour_only(current_minute)
        time_anchor.fill = styles["time_fill"]
        time_anchor.font = styles["header_font"]
        time_anchor.alignment = styles["center"]
        time_anchor.border = styles["thin_border"]

        current_minute += 60
        hour_row_start += slots_per_hour

    ws.freeze_panes = "B7"


def _normalize_breaks_by_shift(breaks_by_shift=None):
    if not breaks_by_shift:
        return {"am": [], "pm": []}
    return {
        "am": list(breaks_by_shift.get("am", [])),
        "pm": list(breaks_by_shift.get("pm", [])),
    }


def create_schedule_workbook(best_answer_text: str, main_rules_path: str, data_path: str, output_path: str, patch_path: str = "", breaks_by_shift=None) -> str:
    clingo = parse_clingo_text(best_answer_text)
    main_rules = parse_lp_file(main_rules_path)
    data_file = parse_lp_file(data_path)
    meta = merge_metadata(main_rules, data_file)
    if patch_path and Path(patch_path).exists():
        meta = merge_metadata(meta, parse_lp_file(patch_path))

    tech_meta = build_tech_metadata(meta)
    techs = list(tech_meta.keys())
    days = get_days(clingo)

    if not techs:
        raise ValueError("No technicians found in tech(...) facts.")
    if not days:
        raise ValueError("No scheduled days found in the best optimisation result.")

    normalized_breaks = _normalize_breaks_by_shift(breaks_by_shift)
    day_schedules = build_day_lane_schedule(clingo, techs, days, tech_meta, normalized_breaks)

    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    for day in days:
        day_name = get_day_name(day, meta)
        short_day = SHORT_DAY_NAMES.get(day_name, day_name[:3])
        weekend = is_weekend(day_name)

        for shift_name in ["am", "pm"]:
            if weekend and shift_name == "pm":
                continue

            shift_techs = [t for t in techs if tech_meta[t]["shift"] == shift_name]
            if not shift_techs:
                continue

            jobs_for_shift = day_schedules[day][shift_name]
            styles = build_styles(shift_name)
            breaks = normalized_breaks.get(shift_name, [])

            raw = str(meta["date_day"].get(day, "")).strip()
            match = re.search(r"\d{4}-\d{2}-\d{2}", raw)

            if match:
                dt = datetime.strptime(match.group(), "%Y-%m-%d")
                short_date = f"{dt.month}-{dt.day}"
            else:
                short_date = raw.replace("/", "-").replace("_", "-")

            base_name = f"{short_day} {short_date} {shift_name.upper()}"
            sheet_name = base_name[:31]
            if sheet_name in wb.sheetnames:
                sheet_name = f"{short_day} {short_date} {shift_name.upper()} {day}"[:31]

            ws = wb.create_sheet(title=sheet_name)
            draw_shift_sheet(ws, day, shift_name, meta, tech_meta, jobs_for_shift, styles, breaks)

        if not weekend:
            combined_jobs = day_schedules[day]["am"] + day_schedules[day]["pm"]
            combined_breaks = sorted(
                normalized_breaks.get("am", []) + normalized_breaks.get("pm", []),
                key=lambda item: item["start"],
            )
            combined_sheet_name = f"{short_day} {short_date} AM+PM"[:31]
            if combined_sheet_name in wb.sheetnames:
                combined_sheet_name = f"{short_day} {short_date} AM+PM {day}"[:31]

            ws = wb.create_sheet(title=combined_sheet_name)
            draw_shift_sheet(
                ws,
                day,
                "combined",
                meta,
                tech_meta,
                combined_jobs,
                build_styles("combined"),
                combined_breaks,
                normalized_breaks,
            )

    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_file)
    return str(output_file)


def create_schedule():
    breaks_by_shift = prompt_break_times_by_shift()

    clingo = parse_clingo(CLINGO_FILE)
    main_rules = parse_lp_file(RULES_FILE)
    data_file = parse_lp_file(DATA_FILE)
    meta = merge_metadata(main_rules, data_file)

    tech_meta = build_tech_metadata(meta)
    techs = list(tech_meta.keys())
    days = get_days(clingo)

    if not techs:
        raise ValueError("No technicians found in tech(...) facts.")
    if not days:
        raise ValueError("No scheduled days found.")

    day_schedules = build_day_lane_schedule(clingo, techs, days, tech_meta, breaks_by_shift)

    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    for day in days:
        day_name = get_day_name(day, meta)
        short_day = SHORT_DAY_NAMES.get(day_name, day_name[:3])
        weekend = is_weekend(day_name)

        for shift_name in ["am", "pm"]:
            if weekend and shift_name == "pm":
                continue

            shift_techs = [t for t in techs if tech_meta[t]["shift"] == shift_name]
            if not shift_techs:
                continue

            jobs_for_shift = day_schedules[day][shift_name]
            styles = build_styles(shift_name)
            breaks = breaks_by_shift.get(shift_name, [])

            raw = str(meta["date_day"].get(day, "")).strip()
            match = re.search(r"\d{4}-\d{2}-\d{2}", raw)

            if match:
                dt = datetime.strptime(match.group(), "%Y-%m-%d")
                short_date = f"{dt.month}-{dt.day}"
            else:
                short_date = raw.replace("/", "-").replace("_", "-")

            base_name = f"{short_day} {short_date} {shift_name.upper()}"
            sheet_name = base_name[:31]

            if sheet_name in wb.sheetnames:
                sheet_name = f"{short_day} {short_date} {shift_name.upper()} {day}"[:31]

            ws = wb.create_sheet(title=sheet_name)
            draw_shift_sheet(ws, day, shift_name, meta, tech_meta, jobs_for_shift, styles, breaks)

        if not weekend:
            combined_jobs = day_schedules[day]["am"] + day_schedules[day]["pm"]
            combined_breaks = sorted(
                breaks_by_shift.get("am", []) + breaks_by_shift.get("pm", []),
                key=lambda item: item["start"],
            )
            combined_sheet_name = f"{short_day} {short_date} AM+PM"[:31]

            if combined_sheet_name in wb.sheetnames:
                combined_sheet_name = f"{short_day} {short_date} AM+PM {day}"[:31]

            ws = wb.create_sheet(title=combined_sheet_name)
            draw_shift_sheet(
                ws,
                day,
                "combined",
                meta,
                tech_meta,
                combined_jobs,
                build_styles("combined"),
                combined_breaks,
                breaks_by_shift,
            )

    wb.save(OUTPUT_FILE)

    print("Done.")
    print("Excel file created here:")
    print(os.path.join(os.getcwd(), OUTPUT_FILE))


if __name__ == "__main__":
    create_schedule()
