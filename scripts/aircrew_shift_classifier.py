#!/usr/bin/env python3
"""Report or classify shuttle crew shifts needed for aircrew arrivals.

Examples:
  ./venv/bin/python scripts/aircrew_shift_classifier.py 10:40 15:03 22:38
  ./venv/bin/python scripts/aircrew_shift_classifier.py --times 10:40,15:03,22:38
  ./venv/bin/python scripts/aircrew_shift_classifier.py --basis-start 2026-07-02 --basis-end 2026-07-15 10:40 17:05 21:22
  APP_ENV=development ./venv/bin/python scripts/aircrew_shift_classifier.py
  ./venv/bin/python scripts/aircrew_shift_classifier.py --env development --month 2026-07 --mode both
  ./venv/bin/python scripts/aircrew_shift_classifier.py --format csv > july_aircrew_crew_shifts.csv
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import Iterable

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def _preconfigure_environment(argv: list[str]) -> None:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--env", choices=["production", "development"])
    parser.add_argument("--database-url")
    args, _ = parser.parse_known_args(argv)
    if args.database_url:
        os.environ["DATABASE_URL"] = args.database_url
    if args.env:
        os.environ["APP_ENV"] = args.env


_preconfigure_environment(sys.argv[1:])

from sqlalchemy import select  # noqa: E402

from app import (  # noqa: E402
    AircrewArrival,
    Assignment,
    Employee,
    NEUTRAL_ASSIGNMENT_VALUES,
    Section,
    SessionLocal,
    _resolve_shuttle_variants_for_values,
    _shift_time_points,
)


@dataclass(frozen=True)
class DayResult:
    shift_date: date
    arrivals: dict[str, list[str]]
    current_crew_shifts: list[str]
    suggested_crew_shifts: list[str]
    rule: str


DEFAULT_BASIS_START = date(2026, 7, 2)
DEFAULT_BASIS_END = date(2026, 7, 15)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Classify the crew shift windows needed to cover aircrew arrivals.",
    )
    parser.add_argument("--env", choices=["production", "development"], help="App database environment to read.")
    parser.add_argument("--database-url", help="Explicit SQLAlchemy database URL, e.g. sqlite:////path/to/schedule.dev.db.")
    parser.add_argument("--start", type=_parse_date, help="First date to inspect, inclusive.")
    parser.add_argument("--end", type=_parse_date, help="Last date to inspect, inclusive.")
    parser.add_argument("--month", help="Month to inspect as YYYY-MM. Defaults to 2026-07.")
    parser.add_argument(
        "--basis-start",
        type=_parse_date,
        default=DEFAULT_BASIS_START,
        help="First schedule date used as the classifier basis. Defaults to 2026-07-02.",
    )
    parser.add_argument(
        "--basis-end",
        type=_parse_date,
        default=DEFAULT_BASIS_END,
        help="Last schedule date used as the classifier basis. Defaults to 2026-07-15.",
    )
    parser.add_argument(
        "--times",
        nargs="+",
        help="Aircrew arrival times to classify directly. Accepts HH:MM, HHMM, AM/PM, or comma-separated values.",
    )
    parser.add_argument(
        "--mode",
        choices=["current", "suggest", "both"],
        default="both",
        help="current = current schedule crew shifts, suggest = decision-tree output, both = side by side.",
    )
    parser.add_argument(
        "--format",
        choices=["table", "csv", "json"],
        default="table",
        help="Output format.",
    )
    parser.add_argument(
        "--show-rules",
        action="store_true",
        help="Print the decision tree before the results.",
    )
    parser.add_argument(
        "--include-employees",
        action="store_true",
        help="Include the employee currently assigned to each crew shift in current-mode output.",
    )
    parser.add_argument(
        "arrival_times",
        nargs="*",
        help="Aircrew arrival times to classify directly, e.g. 10:40 15:03 22:38.",
    )
    return parser.parse_args()


def _parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"expected YYYY-MM-DD, got {value!r}") from exc


def month_bounds(value: str | None) -> tuple[date, date]:
    raw = value or "2026-07"
    try:
        year_s, month_s = raw.split("-", 1)
        year = int(year_s)
        month = int(month_s)
        start = date(year, month, 1)
    except ValueError as exc:
        raise SystemExit(f"--month must be YYYY-MM, got {raw!r}") from exc
    if month == 12:
        next_month = date(year + 1, 1, 1)
    else:
        next_month = date(year, month + 1, 1)
    return start, next_month - timedelta(days=1)


def parse_arrival_times(raw: str | None) -> list[str]:
    if not raw:
        return []
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        return []
    if not isinstance(parsed, list):
        return []
    return sorted({str(item).strip() for item in parsed if str(item).strip()})


def normalize_manual_time(value: str) -> str | None:
    raw = str(value or "").strip()
    if not raw:
        return None
    compact = re.sub(r"\s+", "", raw).lower()

    regular = re.fullmatch(r"(\d{1,2})(?::(\d{2}))?(am|pm)", compact)
    if regular:
        hour = int(regular.group(1))
        minute = int(regular.group(2) or "0")
        if not (1 <= hour <= 12 and 0 <= minute <= 59):
            return None
        hour = (hour % 12) + (12 if regular.group(3) == "pm" else 0)
        return f"{hour:02d}:{minute:02d}"

    military = re.fullmatch(r"(?:(\d{1,2}):(\d{2})|(\d{3,4}))", compact)
    if not military:
        return None
    if military.group(3):
        hour = int(military.group(3)[:-2])
        minute = int(military.group(3)[-2:])
    else:
        hour = int(military.group(1))
        minute = int(military.group(2))
    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        return None
    return f"{hour:02d}:{minute:02d}"


def collect_manual_times(args: argparse.Namespace) -> list[str]:
    raw_tokens = [*(args.times or []), *(args.arrival_times or [])]
    split_tokens: list[str] = []
    for token in raw_tokens:
        split_tokens.extend(part for part in str(token).split(",") if part.strip())

    normalized: list[str] = []
    invalid: list[str] = []
    for token in split_tokens:
        parsed = normalize_manual_time(token)
        if parsed is None:
            invalid.append(token)
        elif parsed not in normalized:
            normalized.append(parsed)
    if invalid:
        raise SystemExit(f"Invalid arrival time(s): {', '.join(invalid)}. Use HH:MM, HHMM, or 3:03pm.")
    return sorted(normalized)


def minutes_from_hhmm(value: str) -> int | None:
    parts = value.strip().split(":")
    if len(parts) != 2:
        return None
    try:
        hour = int(parts[0])
        minute = int(parts[1])
    except ValueError:
        return None
    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        return None
    return hour * 60 + minute


def format_shift_list(values: Iterable[str]) -> str:
    unique = []
    seen = set()
    for value in values:
        normalized = value.strip()
        if normalized and normalized not in seen:
            seen.add(normalized)
            unique.append(normalized)
    return "; ".join(unique) if unique else "-"


def flatten_arrival_minutes(arrivals: dict[str, list[str]]) -> list[int]:
    minutes: list[int] = []
    for times in arrivals.values():
        for value in times:
            parsed = minutes_from_hhmm(value)
            if parsed is not None:
                minutes.append(parsed)
    return sorted(set(minutes))


def decision_tree_text(basis_start: date = DEFAULT_BASIS_START, basis_end: date = DEFAULT_BASIS_END) -> str:
    return "\n".join(
        [
            "Classifier:",
            f"1. Use the current schedule from {basis_start.isoformat()} through {basis_end.isoformat()} as examples.",
            "2. Compare typed arrival times to each example day.",
            "3. Pick the closest example by pickup-time distance, latest evening arrival, and pickup count.",
            "4. Print the crew shifts that were scheduled on that closest July example.",
        ]
    )


def suggest_crew_shifts(arrivals: dict[str, list[str]]) -> tuple[list[str], str]:
    minutes = flatten_arrival_minutes(arrivals)
    if not minutes:
        return [], "no arrivals"

    suggestions: list[str] = []
    day_arrivals = [minute for minute in minutes if 9 * 60 <= minute < 12 * 60]
    evening_arrivals = [minute for minute in minutes if minute >= 14 * 60]

    if day_arrivals:
        suggestions.append("Crew (10:00AM-6:00PM)")

    evening_rule = "no evening arrivals"
    if evening_arrivals:
        latest = max(evening_arrivals)
        if latest >= 23 * 60:
            evening_shift = "4:30pm - 12:30am"
            evening_rule = "latest evening arrival >= 23:00"
        elif latest >= 22 * 60:
            evening_shift = "4:15pm - 12:15am"
            evening_rule = "latest evening arrival >= 22:00"
        elif latest >= (18 * 60) + 30:
            evening_shift = "4:15pm - 11:15pm"
            evening_rule = "latest evening arrival >= 18:30"
        else:
            evening_shift = "4:15pm - 8:15pm"
            evening_rule = "latest evening arrival before 18:30"
        suggestions.append(evening_shift)

    rule_bits = []
    if day_arrivals:
        rule_bits.append("day arrival between 09:00 and 11:59")
    rule_bits.append(evening_rule)
    return suggestions, "; ".join(rule_bits)


def nearest_minute_distance(minutes: list[int], target: int) -> int:
    if not minutes:
        return 24 * 60
    return min(abs(target - minute) for minute in minutes)


def latest_evening_arrival(minutes: list[int]) -> int | None:
    evening = [minute for minute in minutes if minute >= 14 * 60]
    return max(evening) if evening else None


def arrival_pattern_distance(candidate: list[int], example: list[int]) -> int:
    if not candidate and not example:
        return 0
    if not candidate or not example:
        return 10_000

    forward = sum(nearest_minute_distance(example, minute) for minute in candidate)
    backward = sum(nearest_minute_distance(candidate, minute) for minute in example)
    count_penalty = abs(len(candidate) - len(example)) * 90

    candidate_latest = latest_evening_arrival(candidate)
    example_latest = latest_evening_arrival(example)
    if candidate_latest is None and example_latest is None:
        latest_penalty = 0
    elif candidate_latest is None or example_latest is None:
        latest_penalty = 240
    else:
        latest_penalty = abs(candidate_latest - example_latest) * 2

    candidate_has_day = any(9 * 60 <= minute < 12 * 60 for minute in candidate)
    example_has_day = any(9 * 60 <= minute < 12 * 60 for minute in example)
    day_penalty = 240 if candidate_has_day != example_has_day else 0

    return forward + backward + count_penalty + latest_penalty + day_penalty


def load_arrivals(start: date, end: date) -> dict[date, dict[str, list[str]]]:
    with SessionLocal() as session:
        rows = session.scalars(
            select(AircrewArrival)
            .where(AircrewArrival.date >= start, AircrewArrival.date <= end)
            .order_by(AircrewArrival.date, AircrewArrival.carrier)
        ).all()

    by_date: dict[date, dict[str, list[str]]] = {}
    for row in rows:
        by_date.setdefault(row.date, {})[row.carrier] = parse_arrival_times(row.times)
    return by_date


def load_current_crew_shifts(
    start: date,
    end: date,
    arrivals: dict[date, dict[str, list[str]]],
    *,
    include_employees: bool = False,
) -> dict[date, list[str]]:
    with SessionLocal() as session:
        rows = session.execute(
            select(Assignment.date, Employee.name, Assignment.value)
            .join(Employee, Employee.id == Assignment.employee_id)
            .join(Section, Section.id == Employee.section_id)
            .where(
                Section.name == "Shuttle",
                Assignment.date >= start,
                Assignment.date <= end,
            )
            .order_by(Assignment.date, Employee.sort_order.is_(None), Employee.sort_order, Employee.name)
        ).all()

    values_by_date: dict[date, list[tuple[str, str]]] = {}
    for shift_date, employee_name, value in rows:
        values_by_date.setdefault(shift_date, []).append((employee_name, value or ""))

    current: dict[date, list[str]] = {}
    for shift_date, employee_values in values_by_date.items():
        values = [value for _, value in employee_values]
        arrival_minutes = flatten_arrival_minutes(arrivals.get(shift_date, {}))
        variants = _resolve_shuttle_variants_for_values(values, aircrew_arrival_minutes=arrival_minutes)
        crew_values: list[str] = []
        for (employee_name, value), variant in zip(employee_values, variants):
            if variant != "Crew":
                continue
            if not value or value in NEUTRAL_ASSIGNMENT_VALUES:
                continue
            if not _shift_time_points(value):
                continue
            crew_values.append(f"{value} [{employee_name}]" if include_employees else value)
        current[shift_date] = crew_values
    return current


def build_results(start: date, end: date, *, include_employees: bool = False) -> list[DayResult]:
    arrivals = load_arrivals(start, end)
    current = load_current_crew_shifts(
        start,
        end,
        arrivals,
        include_employees=include_employees,
    )
    results: list[DayResult] = []
    day = start
    while day <= end:
        day_arrivals = arrivals.get(day, {})
        suggested, rule = suggest_crew_shifts(day_arrivals)
        results.append(
            DayResult(
                shift_date=day,
                arrivals=day_arrivals,
                current_crew_shifts=current.get(day, []),
                suggested_crew_shifts=suggested,
                rule=rule,
            )
        )
        day += timedelta(days=1)
    return results


def classify_from_basis(
    arrivals: dict[str, list[str]],
    *,
    basis_start: date,
    basis_end: date,
) -> tuple[list[str], str]:
    candidate_minutes = flatten_arrival_minutes(arrivals)
    if not candidate_minutes:
        return [], "no arrivals"

    basis_results = build_results(basis_start, basis_end)
    examples = [
        result
        for result in basis_results
        if result.current_crew_shifts and flatten_arrival_minutes(result.arrivals)
    ]
    if not examples:
        suggestions, rule = suggest_crew_shifts(arrivals)
        return suggestions, f"no usable basis examples from {basis_start} to {basis_end}; fallback rule: {rule}"

    best = min(
        examples,
        key=lambda result: arrival_pattern_distance(
            candidate_minutes,
            flatten_arrival_minutes(result.arrivals),
        ),
    )
    score = arrival_pattern_distance(candidate_minutes, flatten_arrival_minutes(best.arrivals))
    return (
        best.current_crew_shifts,
        f"closest July basis day {best.shift_date.isoformat()} (score {score}): {arrivals_display(best.arrivals)}",
    )


def build_manual_result(times: list[str], *, basis_start: date, basis_end: date) -> list[DayResult]:
    arrivals = {"Manual": times}
    suggested, rule = classify_from_basis(arrivals, basis_start=basis_start, basis_end=basis_end)
    return [
        DayResult(
            shift_date=date.today(),
            arrivals=arrivals,
            current_crew_shifts=[],
            suggested_crew_shifts=suggested,
            rule=rule,
        )
    ]


def arrivals_display(arrivals: dict[str, list[str]]) -> str:
    if not arrivals:
        return "-"
    return "; ".join(f"{carrier}: {'/'.join(times) if times else '-'}" for carrier, times in arrivals.items())


def print_table(results: list[DayResult], mode: str) -> None:
    columns = ["date", "arrivals"]
    if mode in {"current", "both"}:
        columns.append("current crew shifts")
    if mode in {"suggest", "both"}:
        columns.append("suggested crew shifts")
        columns.append("rule")

    rows: list[list[str]] = []
    for result in results:
        row = [result.shift_date.isoformat(), arrivals_display(result.arrivals)]
        if mode in {"current", "both"}:
            row.append(format_shift_list(result.current_crew_shifts))
        if mode in {"suggest", "both"}:
            row.append(format_shift_list(result.suggested_crew_shifts))
            row.append(result.rule)
        rows.append(row)

    widths = [len(column) for column in columns]
    for row in rows:
        for idx, value in enumerate(row):
            widths[idx] = max(widths[idx], len(value))

    print("  ".join(column.ljust(widths[idx]) for idx, column in enumerate(columns)))
    print("  ".join("-" * width for width in widths))
    for row in rows:
        print("  ".join(value.ljust(widths[idx]) for idx, value in enumerate(row)))


def print_csv(results: list[DayResult], mode: str) -> None:
    fieldnames = ["date", "arrivals"]
    if mode in {"current", "both"}:
        fieldnames.append("current_crew_shifts")
    if mode in {"suggest", "both"}:
        fieldnames.extend(["suggested_crew_shifts", "rule"])
    writer = csv.DictWriter(sys.stdout, fieldnames=fieldnames)
    writer.writeheader()
    for result in results:
        row = {
            "date": result.shift_date.isoformat(),
            "arrivals": arrivals_display(result.arrivals),
        }
        if mode in {"current", "both"}:
            row["current_crew_shifts"] = format_shift_list(result.current_crew_shifts)
        if mode in {"suggest", "both"}:
            row["suggested_crew_shifts"] = format_shift_list(result.suggested_crew_shifts)
            row["rule"] = result.rule
        writer.writerow(row)


def print_json(results: list[DayResult], mode: str) -> None:
    payload = []
    for result in results:
        item = {
            "date": result.shift_date.isoformat(),
            "arrivals": result.arrivals,
        }
        if mode in {"current", "both"}:
            item["current_crew_shifts"] = result.current_crew_shifts
        if mode in {"suggest", "both"}:
            item["suggested_crew_shifts"] = result.suggested_crew_shifts
            item["rule"] = result.rule
        payload.append(item)
    print(json.dumps(payload, indent=2))


def print_manual_summary(result: DayResult) -> None:
    print(f"Arrival times: {arrivals_display(result.arrivals).replace('Manual: ', '')}")
    print("Shifts needed:")
    if result.suggested_crew_shifts:
        for shift in result.suggested_crew_shifts:
            print(f"- {shift}")
    else:
        print("- none")
    print(f"Rule: {result.rule}")


def main() -> int:
    args = parse_args()
    manual_times = collect_manual_times(args)
    mode = args.mode
    mode_was_explicit = any(arg == "--mode" or arg.startswith("--mode=") for arg in sys.argv[1:])
    if manual_times and not mode_was_explicit:
        mode = "suggest"
    if manual_times and mode == "current":
        raise SystemExit("--mode current needs database schedule data. Use --mode suggest for typed arrival times.")

    default_start, default_end = month_bounds(args.month)
    start = args.start or default_start
    end = args.end or default_end
    if end < start:
        raise SystemExit("--end must be on or after --start")
    if args.basis_end < args.basis_start:
        raise SystemExit("--basis-end must be on or after --basis-start")

    if args.show_rules:
        print(decision_tree_text(args.basis_start, args.basis_end))
        print()

    if manual_times:
        results = build_manual_result(
            manual_times,
            basis_start=args.basis_start,
            basis_end=args.basis_end,
        )
    else:
        results = build_results(start, end, include_employees=args.include_employees)
    if args.format == "csv":
        print_csv(results, mode)
    elif args.format == "json":
        print_json(results, mode)
    elif manual_times and mode == "suggest":
        print_manual_summary(results[0])
    else:
        print_table(results, mode)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
