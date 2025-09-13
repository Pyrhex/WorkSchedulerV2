from __future__ import annotations

from datetime import date, timedelta
from random import choice, sample
from typing import Dict, List, Optional
import io

from flask import Flask, jsonify, redirect, render_template, request, url_for, send_file, Response
import re
from sqlalchemy import Boolean, Date, ForeignKey, Integer, String, create_engine, func, select, update, delete
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column, relationship, sessionmaker
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


app = Flask(__name__)


@app.after_request
def add_no_cache_headers(resp):
    # Ensure freshly generated schedules show updated coverage highlights immediately
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    resp.headers["Expires"] = "0"
    return resp


# ---- SQLAlchemy setup ----
class Base(DeclarativeBase):
    pass


class Section(Base):
    __tablename__ = "sections"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String, unique=True)
    required_per_day: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    employees: Mapped[List["Employee"]] = relationship("Employee", back_populates="section")


class Employee(Base):
    __tablename__ = "employees"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String, unique=True)
    section_id: Mapped[int] = mapped_column(ForeignKey("sections.id"))
    section: Mapped[Section] = relationship("Section", back_populates="employees")
    # New editable fields
    availability: Mapped[Optional[str]] = mapped_column(String, nullable=True)  # freeform notes or JSON
    preferred_shift: Mapped[Optional[str]] = mapped_column(String, nullable=True)
    seniority: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    preferred_shifts_per_week: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)
    max_shifts_per_week: Mapped[Optional[int]] = mapped_column(Integer, nullable=True)


class EmployeeRole(Base):
    __tablename__ = "employee_roles"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    employee_id: Mapped[int] = mapped_column(ForeignKey("employees.id"))
    section_id: Mapped[int] = mapped_column(ForeignKey("sections.id"))


class EmployeeAvailability(Base):
    __tablename__ = "employee_availability"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    employee_id: Mapped[int] = mapped_column(ForeignKey("employees.id"))
    day_of_week: Mapped[int] = mapped_column(Integer)  # 0=Mon .. 6=Sun
    shift_label: Mapped[str] = mapped_column(String)   # matches one of role shift variants
    allowed: Mapped[bool] = mapped_column(Boolean, default=True)


class Week(Base):
    __tablename__ = "weeks"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    start_date: Mapped[date] = mapped_column(Date, unique=True)  # Thu start


class Assignment(Base):
    __tablename__ = "assignments"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    week_id: Mapped[int] = mapped_column(ForeignKey("weeks.id"))
    employee_id: Mapped[int] = mapped_column(ForeignKey("employees.id"))
    date: Mapped[date] = mapped_column(Date)
    value: Mapped[str] = mapped_column(String)  # shift label, e.g., "Set", "6AM–2PM"
    # Mark when the scheduler would have assigned a shift but skipped due to time off
    dismissed_timeoff: Mapped[bool] = mapped_column(Boolean, default=False)


class TimeOff(Base):
    __tablename__ = "timeoff"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String)
    role: Mapped[str] = mapped_column(String)
    from_date: Mapped[date] = mapped_column(Date)
    to_date: Mapped[date] = mapped_column(Date)
    approved: Mapped[bool] = mapped_column(Boolean, default=False)
    vacation: Mapped[bool] = mapped_column(Boolean, default=False)


engine = create_engine("sqlite:///schedule.db", future=True)
SessionLocal = sessionmaker(bind=engine, future=True, expire_on_commit=False)


# ---- Helpers and constants ----
def daterange(start: date, days: int) -> List[date]:
    return [start + timedelta(days=i) for i in range(days)]


def week_dates(start: date) -> List[dict]:
    result = []
    for d in daterange(start, 7):
        result.append(
            {
                "key": d.isoformat(),
                "label_short": d.strftime("%a %m/%d"),
                "label_long": d.strftime("%a %b %d, %Y"),
            }
        )
    return result


def format_week_label(start: date) -> str:
    end = start + timedelta(days=6)
    return f"{start.strftime('%a %b %d, %Y')} – {end.strftime('%a %b %d, %Y')}"


TIME_OFF_LABEL = "TIME OFF"

# Seniority order for Front Desk manager-on-duty selection
SENIORITY_ORDER = [
    "Cindy", "KC", "Ryan", "Emilyn", "Christian", "Troy",
    "Brian", "Tristan", "Terry", "Jordan", "Abdi", "Sato"
]

BREAKFAST_SHIFTS = [
    "Set",
    TIME_OFF_LABEL,
    "5AM–12PM",
    "6AM–12PM",
    "7AM–12PM",
]

# Front Desk: three variants (AM, PM, Audit), each has two staggered times
FRONT_DESK_SHIFTS = [
    "Set",
    TIME_OFF_LABEL,
    "AM (6:00AM–2:00PM)",
    "AM (6:15AM–2:15PM)",
    "PM (2:00PM–10:00PM)",
    "PM (2:15PM–10:15PM)",
    "Audit (10:00PM–6:00AM)",
    "Audit (10:15PM–6:15AM)",
]

# Shuttle: four fixed 8-hour variants
SHUTTLE_SHIFTS = [
    "Set",
    TIME_OFF_LABEL,
    "AM (3:30AM–11:30AM)",
    "Midday (10:30AM–6:30PM)",
    "PM (5:30PM–1:30AM)",
    "Crew (5:45PM–1:45AM)",
]


def init_db_once():
    Base.metadata.create_all(engine)
    with SessionLocal() as s:
        # Ensure new columns exist for employees (SQLite light migration)
        with engine.connect() as conn:
            cols = [row[1] for row in conn.exec_driver_sql("PRAGMA table_info(employees)").fetchall()]
            if "availability" not in cols:
                conn.exec_driver_sql("ALTER TABLE employees ADD COLUMN availability TEXT")
            if "preferred_shift" not in cols:
                conn.exec_driver_sql("ALTER TABLE employees ADD COLUMN preferred_shift TEXT")
            if "seniority" not in cols:
                conn.exec_driver_sql("ALTER TABLE employees ADD COLUMN seniority INTEGER")
            to_cols = [row[1] for row in conn.exec_driver_sql("PRAGMA table_info(timeoff)").fetchall()]
            if "vacation" not in to_cols:
                conn.exec_driver_sql("ALTER TABLE timeoff ADD COLUMN vacation INTEGER DEFAULT 0")
            if "preferred_shifts_per_week" not in cols:
                conn.exec_driver_sql("ALTER TABLE employees ADD COLUMN preferred_shifts_per_week INTEGER")
            if "max_shifts_per_week" not in cols:
                conn.exec_driver_sql("ALTER TABLE employees ADD COLUMN max_shifts_per_week INTEGER")
            # Create employee_roles table if missing
            conn.exec_driver_sql(
                """
                CREATE TABLE IF NOT EXISTS employee_roles (
                    id INTEGER PRIMARY KEY,
                    employee_id INTEGER NOT NULL,
                    section_id INTEGER NOT NULL,
                    FOREIGN KEY(employee_id) REFERENCES employees(id),
                    FOREIGN KEY(section_id) REFERENCES sections(id)
                )
                """
            )
        # Ensure new column exists for assignments.dismissed_timeoff (SQLite light migration)
        with engine.connect() as conn:
            a_cols = [row[1] for row in conn.exec_driver_sql("PRAGMA table_info(assignments)").fetchall()]
            if "dismissed_timeoff" not in a_cols:
                conn.exec_driver_sql("ALTER TABLE assignments ADD COLUMN dismissed_timeoff INTEGER DEFAULT 0")

        # Ensure sections exist (and update FD required to 6)
        names = {n: None for n in ["Breakfast Bar", "Front Desk", "Shuttle"]}
        existing = {sec.name: sec for sec in s.scalars(select(Section))}
        for n in names.keys():
            if n in existing:
                if n == "Front Desk" and (existing[n].required_per_day or 0) != 6:
                    existing[n].required_per_day = 6
            else:
                rp = 6 if n == "Front Desk" else None
                s.add(Section(name=n, required_per_day=rp))
        s.flush()

        # Refresh sections after potential inserts
        sections = {sec.name: sec for sec in s.scalars(select(Section))}

        # Ensure baseline employees exist
        # def ensure_emp(name: str, sec_name: str):
        #     e = s.scalar(select(Employee).where(Employee.name == name))
        #     if not e:
        #         s.add(Employee(name=name, section_id=sections[sec_name].id))

        # for n in ["Rose", "Yoko", "Eurielle", "Anna", "Merve", "Ayako"]:
        #     ensure_emp(n, "Breakfast Bar")
        # for n in ["Ryan", "Emilyn", "Abdi", "Jordan", "Cindy"]:
        #     ensure_emp(n, "Front Desk")
        # for n in ["Alex", "Taylor", "Morgan"]:
        #     ensure_emp(n, "Shuttle")

        # # Seed time off if empty
        # if not s.scalar(select(TimeOff).limit(1)):
        #     to_items = [
        #         TimeOff(name="Christian", role="Front Desk", from_date=date(2025, 9, 10), to_date=date(2025, 9, 24), approved=True),
        #         TimeOff(name="Abdi", role="Front Desk", from_date=date(2025, 9, 18), to_date=date(2025, 9, 30), approved=True),
        #         TimeOff(name="Cindy", role="Front Desk", from_date=date(2025, 9, 18), to_date=date(2025, 9, 18), approved=False),
        #         TimeOff(name="Ayako", role="Breakfast Bar", from_date=date(2025, 9, 19), to_date=date(2025, 9, 19), approved=False),
        #         TimeOff(name="Merve", role="Breakfast Bar", from_date=date(2025, 9, 19), to_date=date(2025, 9, 19), approved=False),
        #         TimeOff(name="Tristan", role="Front Desk", from_date=date(2025, 9, 19), to_date=date(2025, 9, 20), approved=True),
        #     ]
        #     s.add_all(to_items)

        # Ensure there is a baseline week
        wk = s.scalar(select(Week).where(Week.start_date == date(2025, 9, 18)))
        if not wk:
            wk = Week(start_date=date(2025, 9, 18))
            s.add(wk)
            s.flush()

        # Ensure assignments exist for every employee/date
        day_list = daterange(wk.start_date, 7)
        for emp in s.scalars(select(Employee)):
            for d in day_list:
                exists = s.scalar(select(Assignment).where(Assignment.week_id == wk.id, Assignment.employee_id == emp.id, Assignment.date == d))
                if not exists:
                    s.add(Assignment(week_id=wk.id, employee_id=emp.id, date=d, value="Set"))

        s.commit()

        # Ensure TIME OFF is synchronized into assignments for approved requests
        sync_timeoff_to_assignments(wk.id, s)

        # Seed sample assignments if week is all Set
        any_non_set = s.scalar(select(Assignment).where(Assignment.week_id == wk.id, Assignment.value != "Set").limit(1))
        if not any_non_set:
            seed_example_assignments_db(wk.id, s)
            s.commit()


def has_approved_timeoff(name: str, dte: date, s: Session) -> bool:
    to = s.scalar(
        select(TimeOff).where(
            TimeOff.name == name,
            TimeOff.approved.is_(True),
            TimeOff.from_date <= dte,
            TimeOff.to_date >= dte,
        )
    )
    return to is not None

def has_any_timeoff(name: str, dte: date, s: Session) -> bool:
    to = s.scalar(
        select(TimeOff).where(
            TimeOff.name == name,
            TimeOff.from_date <= dte,
            TimeOff.to_date >= dte,
        )
    )
    return to is not None


def sync_timeoff_to_assignments(week_id: int, s: Session):
    wk = s.get(Week, week_id)
    days = list(daterange(wk.start_date, 7))
    # For each employee and day, if approved time off, set TIME OFF label
    for emp in s.scalars(select(Employee)):
        for d in days:
            if has_approved_timeoff(emp.name, d, s):
                a = s.scalar(select(Assignment).where(Assignment.week_id == week_id, Assignment.employee_id == emp.id, Assignment.date == d))
                if a and a.value != TIME_OFF_LABEL:
                    a.value = TIME_OFF_LABEL
    s.commit()


def seed_example_assignments_db(week_id: int, s: Session):
    # Breakfast Bar, Front Desk, Shuttle examples
    bb_emp_ids = [e.id for e in s.scalars(select(Employee).join(Section).where(Section.name == "Breakfast Bar"))]
    fd_emp_ids = [e.id for e in s.scalars(select(Employee).join(Section).where(Section.name == "Front Desk"))]
    sh_emp_ids = [e.id for e in s.scalars(select(Employee).join(Section).where(Section.name == "Shuttle"))]
    wk = s.get(Week, week_id)
    for d in daterange(wk.start_date, 7):
        # Breakfast Bar
        for eid in bb_emp_ids[:2]:  # only a couple for color variety
            emp = s.get(Employee, eid)
            if emp and has_any_timeoff(emp.name, d, s):
                continue
            a = s.scalar(select(Assignment).where(Assignment.week_id == week_id, Assignment.employee_id == eid, Assignment.date == d))
            if a:
                a.value = choice(["5AM–12PM", "6AM–12PM", "7AM–12PM", "Set"])  # greens/blues
        # Front Desk: sample among AM/PM/Audit staggered options
        for eid in fd_emp_ids:
            emp = s.get(Employee, eid)
            if emp and has_any_timeoff(emp.name, d, s):
                continue
            a = s.scalar(select(Assignment).where(Assignment.week_id == week_id, Assignment.employee_id == eid, Assignment.date == d))
            if a:
                a.value = choice([
                    "AM (6:00AM–2:00PM)", "AM (6:15AM–2:15PM)",
                    "PM (2:00PM–10:00PM)", "PM (2:15PM–10:15PM)",
                    "Audit (10:00PM–6:00AM)", "Audit (10:15PM–6:15AM)",
                    "Set"
                ])

        # Shuttle
        for eid in sh_emp_ids:
            emp = s.get(Employee, eid)
            if emp and has_any_timeoff(emp.name, d, s):
                continue
            a = s.scalar(select(Assignment).where(Assignment.week_id == week_id, Assignment.employee_id == eid, Assignment.date == d))
            if a:
                a.value = choice([
                    "AM (3:30AM–11:30AM)",
                    "Midday (10:30AM–6:30PM)",
                    "PM (5:30PM–1:30AM)",
                    "Crew (5:45PM–1:45AM)",
                    "Set"
                ])


def has_generated_schedule(week_id: int) -> bool:
    """Check if the schedule has already been generated (has non-'Set' assignments)"""
    with SessionLocal() as s:
        any_non_set = s.scalar(select(Assignment).where(Assignment.week_id == week_id, Assignment.value != "Set").limit(1))
        return any_non_set is not None


def coverage_snapshot_db(week_id: int) -> tuple[dict, dict, int, dict, dict, int, dict, dict, int, dict, dict]:
    with SessionLocal() as s:
        wk = s.get(Week, week_id)
        dates = [d.isoformat() for d in daterange(wk.start_date, 7)]
        
        # Initialize Front Desk counts (2 per variant required)
        shift_variants = ["AM", "PM", "Audit"]
        counts = {k: {variant: 0 for variant in shift_variants} for k in dates}
        missing = {k: False for k in dates}

        # Initialize Shuttle counts (1 per variant required)
        sh_variants = ["AM", "Midday", "PM", "Crew"]
        sh_counts = {k: {variant: 0 for variant in sh_variants} for k in dates}
        sh_missing = {k: False for k in dates}

        # Initialize Breakfast counts (1 per variant required)
        bb_variants = ["5AM–12PM", "6AM–12PM", "7AM–12PM"]
        bb_counts = {k: {variant: 0 for variant in bb_variants} for k in dates}
        bb_missing = {k: False for k in dates}

        # Track Front Desk duplicates of exact staggered times per variant (e.g., two at 2:15PM)
        fd_duplicates = {k: False for k in dates}
        # Exact label counts per date and variant for Front Desk
        fd_label_counts: dict[str, dict[str, dict[str, int]]] = {k: {"AM": {}, "PM": {}, "Audit": {}} for k in dates}
        
        # Count Front Desk assignments per shift variant per day
        # Include any employee assigned to a Front Desk-like label (AM/PM/Audit),
        # regardless of primary role, to account for secondary-role coverage.
        rows = s.scalars(select(Assignment).where(Assignment.week_id == week_id))

        for a in rows:
            if not a.value or a.value in ("Set", TIME_OFF_LABEL):
                continue

            date_key = a.date.isoformat()

            # Front Desk variants: count ONLY if the value is a Front Desk label
            if a.value in FRONT_DESK_SHIFTS:
                if a.value.startswith("AM"):
                    counts[date_key]["AM"] += 1
                    fd_label_counts[date_key]["AM"][a.value] = fd_label_counts[date_key]["AM"].get(a.value, 0) + 1
                elif a.value.startswith("PM"):
                    counts[date_key]["PM"] += 1
                    fd_label_counts[date_key]["PM"][a.value] = fd_label_counts[date_key]["PM"].get(a.value, 0) + 1
                elif a.value.startswith("Audit"):
                    counts[date_key]["Audit"] += 1
                    fd_label_counts[date_key]["Audit"][a.value] = fd_label_counts[date_key]["Audit"].get(a.value, 0) + 1

            # Shuttle variants: count ONLY if the value is a Shuttle label
            if a.value in SHUTTLE_SHIFTS:
                if a.value == "AM (3:30AM–11:30AM)":
                    sh_counts[date_key]["AM"] += 1
                elif a.value.startswith("Midday"):
                    sh_counts[date_key]["Midday"] += 1
                elif a.value.startswith("PM (5:30PM"):
                    sh_counts[date_key]["PM"] += 1
                elif a.value.startswith("Crew"):
                    sh_counts[date_key]["Crew"] += 1

            # Breakfast variants (exact labels)
            if a.value in bb_variants:
                bb_counts[date_key][a.value] += 1
        
        # Check for missing coverage: each variant needs at least 2 people
        for date_key in dates:
            for variant in shift_variants:
                if counts[date_key][variant] < 2:
                    missing[date_key] = True
                    break

        # Compute duplicate-stagger warnings for Front Desk per date
        for date_key in dates:
            dup_any = False
            for variant in ("AM", "PM", "Audit"):
                exact = fd_label_counts[date_key][variant]
                total = sum(exact.values())
                distinct = sum(1 for v in exact.values() if v > 0)
                if total >= 2 and distinct == 1:
                    dup_any = True
                    break
            fd_duplicates[date_key] = dup_any
        
        # For backward compatibility, also return total counts
        total_counts = {k: sum(counts[k].values()) for k in dates}
        required = 6  # 2 per variant * 3 variants
        # Shuttle missing and required (1 per each of 4 variants)
        for date_key in dates:
            for variant in sh_variants:
                if sh_counts[date_key][variant] < 1:
                    sh_missing[date_key] = True
                    break
        sh_required = 4

        # Breakfast missing and required (1 per each of 3 variants)
        for date_key in dates:
            for variant in bb_variants:
                if bb_counts[date_key][variant] < 1:
                    bb_missing[date_key] = True
                    break
        bb_required = 3

        return total_counts, missing, required, counts, sh_missing, sh_required, sh_counts, bb_missing, bb_required, bb_counts, fd_duplicates


def double_booked_snapshot(week_id: int) -> dict[str, list[str]]:
    """Return date_key -> list of employee names double-booked (active in 2+ sections same day)."""
    with SessionLocal() as s:
        sections = list(s.scalars(select(Section)))
        sec_by_id = {sec.id: sec.name for sec in sections}
        emp_sections: dict[int, set[str]] = {}
        for e in s.scalars(select(Employee)):
            emp_sections.setdefault(e.id, set()).add(sec_by_id.get(e.section_id, ""))
        for er in s.scalars(select(EmployeeRole)):
            emp_sections.setdefault(er.employee_id, set()).add(sec_by_id.get(er.section_id, ""))

        multi_emp_ids = {eid for eid, secset in emp_sections.items() if len([n for n in secset if n]) > 1}
        if not multi_emp_ids:
            return {}

        def section_shifts(name: str) -> list[str]:
            if name == "Breakfast Bar":
                return BREAKFAST_SHIFTS
            if name == "Front Desk":
                return FRONT_DESK_SHIFTS
            if name == "Shuttle":
                return SHUTTLE_SHIFTS
            return []

        result: dict[str, list[str]] = {}
        wk = s.get(Week, week_id)
        if not wk:
            return {}
        for eid in multi_emp_ids:
            emp = s.get(Employee, eid)
            if not emp:
                continue
            for d in daterange(wk.start_date, 7):
                a = s.scalar(select(Assignment).where(Assignment.week_id == week_id, Assignment.employee_id == eid, Assignment.date == d))
                if not a or not a.value or a.value in ("Set", TIME_OFF_LABEL):
                    continue
                active = 0
                for sec_name in (emp_sections.get(eid) or set()):
                    if sec_name and a.value in section_shifts(sec_name):
                        active += 1
                if active > 1:
                    key = d.isoformat()
                    result.setdefault(key, []).append(emp.name)
        return result


def build_week_context(week_id: int):
    with SessionLocal() as s:
        wk = s.get(Week, week_id)
        dates = week_dates(wk.start_date)
        week_keys = {d["key"] for d in dates}
        # Sections and employees
        sections = {sec.name: {"employees": [], "assignments": {}, "shifts": []} for sec in s.scalars(select(Section))}
        for sec_name in sections.keys():
            # Include employees with this as primary role and any with this as a secondary role
            sec_obj = s.scalar(select(Section).where(Section.name == sec_name))
            if not sec_obj:
                continue
            # Preserve the same ordering as Employees tab: DB insertion (id asc) for primaries
            primary_emps = list(
                s.scalars(
                    select(Employee).where(Employee.section_id == sec_obj.id).order_by(Employee.id.asc())
                )
            )
            secondary_ids = [row[0] for row in s.execute(select(EmployeeRole.employee_id).where(EmployeeRole.section_id == sec_obj.id)).all()]
            secondary_emps = (
                list(s.scalars(select(Employee).where(Employee.id.in_(secondary_ids)).order_by(Employee.id.asc()))) if secondary_ids else []
            )
            primary_ids = {e.id for e in primary_emps}
            # Append secondary employees after primaries, keep their insertion order, and avoid duplicates
            emps = primary_emps + [e for e in secondary_emps if e.id not in primary_ids]
            sections[sec_name]["employees"] = [e.name for e in emps]
            # init assignment map
            for e in emps:
                sections[sec_name]["assignments"][e.name] = {d["key"]: "Set" for d in dates}
            # fill from DB
            for e in emps:
                rows = s.scalars(select(Assignment).where(Assignment.week_id == week_id, Assignment.employee_id == e.id))
                for a in rows:
                    sections[sec_name]["assignments"][e.name][a.date.isoformat()] = a.value
            # shifts per section
            if sec_name == "Breakfast Bar":
                sections[sec_name]["shifts"] = BREAKFAST_SHIFTS
            elif sec_name == "Front Desk":
                sections[sec_name]["shifts"] = FRONT_DESK_SHIFTS
            elif sec_name == "Shuttle":
                sections[sec_name]["shifts"] = SHUTTLE_SHIFTS

        # Detect employees listed in multiple sections
        emp_in_sections: dict[str, set[str]] = {}
        for sec_name, sec in sections.items():
            for nm in sec["employees"]:
                emp_in_sections.setdefault(nm, set()).add(sec_name)
        multi_role_names = {nm for nm, secset in emp_in_sections.items() if len(secset) > 1}

        # Build double-booked map: date_key -> names who appear to have active shifts in 2+ sections the same day
        def section_shifts(name: str) -> list[str]:
            if name == "Breakfast Bar":
                return BREAKFAST_SHIFTS
            if name == "Front Desk":
                return FRONT_DESK_SHIFTS
            if name == "Shuttle":
                return SHUTTLE_SHIFTS
            return []

        double_booked: dict[str, list[str]] = {d["key"]: [] for d in dates}
        if multi_role_names:
            for nm in multi_role_names:
                for d in dates:
                    active = 0
                    for sec_name in emp_in_sections.get(nm, set()):
                        # Only count if the value belongs to that section's shift options and is not Set/OFF
                        val = sections[sec_name]["assignments"][nm][d["key"]]
                        if val and val not in ("Set", TIME_OFF_LABEL) and val in section_shifts(sec_name):
                            active += 1
                    if active > 1:
                        double_booked[d["key"]].append(nm)

        # time off
        to_list = []
        vacation_days: dict[str, set[str]] = {}
        dismissed_days: dict[str, set[str]] = {}
        for t in s.scalars(select(TimeOff)):
            same_day = t.from_date == t.to_date
            label = t.from_date.strftime("%b %d") if same_day else f"{t.from_date.strftime('%b %d')} to {t.to_date.strftime('%b %d')}"
            to_list.append({
                "id": t.id,
                "name": t.name,
                "role": t.role,
                "label": label,
                "approved": t.approved,
                "vacation": bool(getattr(t, 'vacation', False)),
            })
            # Build per-employee per-day vacation map for current week
            if bool(getattr(t, 'vacation', False)):
                cur = t.from_date
                while cur <= t.to_date:
                    k = cur.isoformat()
                    if k in week_keys:
                        vacation_days.setdefault(t.name, set()).add(k)
                    cur += timedelta(days=1)

        # Build dismissed_days map from assignments flag for this week (guard if column exists)
        if hasattr(Assignment, 'dismissed_timeoff'):
            for a in s.scalars(select(Assignment).where(Assignment.week_id == wk.id, Assignment.dismissed_timeoff == 1)):
                emp = s.get(Employee, a.employee_id)
                if not emp:
                    continue
                dismissed_days.setdefault(emp.name, set()).add(a.date.isoformat())

        # Check if schedule has been generated
        schedule_generated = has_generated_schedule(week_id)
        
        meta = {
            "range_label": "September 18 – October 15, 2025",
            "success_banner": "4-week schedule saved from September 18, 2025 to October 15, 2025!",
            "week_label": f"{format_week_label(wk.start_date)}",
            "fd_note": "Front Desk: 2 agents per AM/PM/Audit (6 total/day)",
            "schedule_generated": schedule_generated,
            "week_id": week_id,
        }

        return {
            "week": {"label": "Week 1", "dates": dates, "sections": sections},
            "breakfast": sections["Breakfast Bar"],
            "front_desk": sections["Front Desk"],
            "meta": meta,
            "time_off": to_list,
            "vacation_days": vacation_days,
            "dismissed_days": dismissed_days,
            "double_booked": double_booked,
        }


# Initialize DB at import time (Flask 3.x removed before_first_request)
init_db_once()

# ---- Routes ----


@app.route("/")
def index():
    with SessionLocal() as s:
        week = s.scalar(select(Week).where(Week.start_date == date(2025, 9, 18)))
        ctx = build_week_context(week.id)
    counts, missing, required, variant_counts, shuttle_missing, shuttle_required, shuttle_counts, bb_missing, bb_required, bb_counts, fd_duplicates = coverage_snapshot_db(week.id)
    return render_template(
        "schedule.html",
        meta=ctx["meta"],
        week=ctx["week"],
        breakfast=ctx["breakfast"],
        front_desk=ctx["front_desk"],
        shuttle=ctx["week"]["sections"].get("Shuttle"),
        counts=counts,
        missing=missing,
        required=required,
        variant_counts=variant_counts,
        shuttle_missing=shuttle_missing,
        shuttle_required=shuttle_required,
        shuttle_counts=shuttle_counts,
        bb_missing=bb_missing,
        bb_required=bb_required,
        bb_counts=bb_counts,
        fd_duplicates=fd_duplicates,
        time_off=ctx["time_off"],
        vacation_days=ctx.get("vacation_days", {}),
        dismissed_days=ctx.get("dismissed_days", {}),
        double_booked=ctx["double_booked"],
    )


@app.template_filter("fd_display")
def fd_display(label: str) -> str:
    # For Front Desk shift labels like "AM (6:00AM–2:00PM)", show only the time range
    if label in ("Set", TIME_OFF_LABEL):
        return label
    if "(" in label and ")" in label:
        start = label.find("(") + 1
        end = label.find(")", start)
        if end > start:
            return label[start:end]
    return label


@app.template_filter("time_only")
def time_only(label: str) -> str:
    # For labels like "Midday (10:30AM–6:30PM)", show only the time range
    if label in ("Set", TIME_OFF_LABEL):
        return label
    if "(" in label and ")" in label:
        start = label.find("(") + 1
        end = label.find(")", start)
        if end > start:
            return label[start:end]
    return label


@app.template_filter("format_shift")
def format_shift(label: str) -> str:
    # Map Set -> "-" for compact display
    if not label:
        return label
    if label.lower() == "set":
        return "-"
    if label == TIME_OFF_LABEL:
        return label
    # Prefer time-only inside parentheses if present
    if "(" in label and ")" in label:
        start = label.find("(") + 1
        end = label.find(")", start)
        if end > start:
            label = label[start:end]
    # Normalize dash spacing around en dash
    label = re.sub(r"\s*–\s*", " – ", label)
    # Also handle simple hyphen just in case
    label = re.sub(r"\s*-\s*", " - ", label)
    # Lowercase AM/PM tokens (handles attached forms like 6:00AM)
    label = re.sub(r"AM|PM", lambda m: m.group(0).lower(), label)
    return label


@app.route("/week/<int:week_id>")
def view_week(week_id: int):
    ctx = build_week_context(week_id)
    counts, missing, required, variant_counts, shuttle_missing, shuttle_required, shuttle_counts, bb_missing, bb_required, bb_counts, fd_duplicates = coverage_snapshot_db(week_id)
    return render_template(
        "schedule.html",
        meta=ctx["meta"],
        week=ctx["week"],
        breakfast=ctx["breakfast"],
        front_desk=ctx["front_desk"],
        shuttle=ctx["week"]["sections"].get("Shuttle"),
        counts=counts,
        missing=missing,
        required=required,
        variant_counts=variant_counts,
        shuttle_missing=shuttle_missing,
        shuttle_required=shuttle_required,
        shuttle_counts=shuttle_counts,
        bb_missing=bb_missing,
        bb_required=bb_required,
        bb_counts=bb_counts,
        fd_duplicates=fd_duplicates,
        time_off=ctx["time_off"],
        vacation_days=ctx.get("vacation_days", {}),
        dismissed_days=ctx.get("dismissed_days", {}),
        double_booked=ctx["double_booked"],
    )


def _fd_variant(label: str) -> Optional[str]:
    if not label:
        return None
    if label.startswith("AM"):
        return "AM"
    if label.startswith("PM"):
        return "PM"
    if label.startswith("Audit"):
        return "Audit"
    return None


@app.route("/week/<int:week_id>/manager-meals")
def manager_meals(week_id: int):
    # Build a map of day -> {variant -> list of employee names on that FD shift}
    with SessionLocal() as s:
        wk = s.get(Week, week_id)
        if not wk:
            return redirect(url_for("view_week", week_id=week_id))

        # Pull all assignments for the week and map by day/variant
        by_day: dict[date, dict[str, list[str]]] = {}
        rows = list(s.scalars(select(Assignment).where(Assignment.week_id == week_id)))
        # Preload employee id -> name for efficiency
        emp_name = {e.id: e.name for e in s.scalars(select(Employee))}
        for a in rows:
            if not a.value or a.value in ("Set", TIME_OFF_LABEL):
                continue
            variant = _fd_variant(a.value)
            if not variant:
                continue
            # Only consider Front Desk-like labels
            if variant not in ("AM", "PM", "Audit"):
                continue
            nm = emp_name.get(a.employee_id)
            if not nm:
                continue
            by_day.setdefault(a.date, {}).setdefault(variant, []).append(nm)

        # Decide manager meal recipient per day/variant using seniority
        # AM Sunday-Thursday override goes to JoJo regardless of schedule
        results: list[dict] = []
        for d in daterange(wk.start_date, 7):
            entry = {
                "date": d,
                "label": d.strftime("%a %m/%d"),
                "AM": None,
                "PM": None,
                "Audit": None,
            }
            # AM override: Sunday (6) through Thursday (3)
            if d.weekday() in (6, 0, 1, 2, 3):
                entry["AM"] = "JoJo"
            else:
                am_list = by_day.get(d, {}).get("AM", [])
                entry["AM"] = next((n for n in SENIORITY_ORDER if n in am_list), (am_list[0] if am_list else None))
            # PM and Audit by seniority among assigned that shift
            for variant in ("PM", "Audit"):
                names = by_day.get(d, {}).get(variant, [])
                chosen = next((n for n in SENIORITY_ORDER if n in names), (names[0] if names else None))
                entry[variant] = chosen
            results.append(entry)

        # Build simple copy-friendly rows: Date -> "Night, AM, PM"
        rows: list[tuple[str, str]] = []
        for e in results:
            date_str = e["date"].strftime("%a %m/%d")
            audit = e.get("Audit") or "-"
            am = e.get("AM") or "-"
            pm = e.get("PM") or "-"
            rows.append((date_str, f"{audit}, {am}, {pm}"))

        return render_template(
            "meals.html",
            rows=rows,
            error=None,
        )


@app.route("/week/<int:week_id>/manager-meals.txt")
def manager_meals_text(week_id: int):
    # Plain-text export in format: Date: Audit, AM, PM
    with SessionLocal() as s:
        wk = s.get(Week, week_id)
        if not wk:
            return redirect(url_for("view_week", week_id=week_id))

        # Build assignments by day/variant
        by_day: dict[date, dict[str, list[str]]] = {}
        rows = list(s.scalars(select(Assignment).where(Assignment.week_id == week_id)))
        emp_name = {e.id: e.name for e in s.scalars(select(Employee))}
        for a in rows:
            if not a.value or a.value in ("Set", TIME_OFF_LABEL):
                continue
            variant = _fd_variant(a.value)
            if not variant:
                continue
            if variant not in ("AM", "PM", "Audit"):
                continue
            nm = emp_name.get(a.employee_id)
            if not nm:
                continue
            by_day.setdefault(a.date, {}).setdefault(variant, []).append(nm)

        lines: list[str] = []
        for d in daterange(wk.start_date, 7):
            # Decide AM
            if d.weekday() in (6, 0, 1, 2, 3):
                am_name = "JoJo"
            else:
                am_list = by_day.get(d, {}).get("AM", [])
                am_name = next((n for n in SENIORITY_ORDER if n in am_list), (am_list[0] if am_list else None))
            # PM and Audit
            def pick(variant: str) -> Optional[str]:
                names = by_day.get(d, {}).get(variant, [])
                return next((n for n in SENIORITY_ORDER if n in names), (names[0] if names else None))
            audit_name = pick("Audit")
            pm_name = pick("PM")
            label = d.strftime("%a %m/%d")
            lines.append(f"{label}: {audit_name or '-'}, {am_name or '-'}, {pm_name or '-'}")

        text = "\n".join(lines)
        return Response(text, mimetype="text/plain")


    


@app.route("/admin/employees/add", methods=["GET", "POST"]) 
def admin_add_employee():
    if request.method == "GET":
        return render_template("admin_add_employee.html")
    # POST
    name = (request.form.get("name") or "").strip()
    role = (request.form.get("role") or "").strip()
    if not name or not role:
        return render_template("admin_add_employee.html", error="Name and role are required"), 400
    with SessionLocal() as s:
        # Check duplicates
        exists = s.scalar(select(Employee).where(Employee.name == name))
        if exists:
            return render_template("admin_add_employee.html", error="Employee with this name already exists"), 400
        sec = s.scalar(select(Section).where(Section.name == role))
        if not sec:
            return render_template("admin_add_employee.html", error="Unknown role"), 400
        # Optional fields
        preferred_shift = (request.form.get("preferred_shift") or "").strip() or None
        seniority_raw = (request.form.get("seniority") or "").strip()
        seniority = int(seniority_raw) if seniority_raw.isdigit() else None
        pref_count_raw = (request.form.get("preferred_shifts_per_week") or "").strip()
        max_count_raw = (request.form.get("max_shifts_per_week") or "").strip()
        preferred_shifts_per_week = int(pref_count_raw) if pref_count_raw.isdigit() else None
        max_shifts_per_week = int(max_count_raw) if max_count_raw.isdigit() else None
        availability = (request.form.get("availability") or "").strip() or None
        emp = Employee(
            name=name,
            section_id=sec.id,
            preferred_shift=preferred_shift,
            seniority=seniority,
            preferred_shifts_per_week=preferred_shifts_per_week,
            max_shifts_per_week=max_shifts_per_week,
            availability=availability,
        )
        s.add(emp)
        s.flush()
        # Ensure assignments for current week
        week = s.scalar(select(Week).where(Week.start_date == date(2025, 9, 18)))
        for d in daterange(week.start_date, 7):
            s.add(Assignment(week_id=week.id, employee_id=emp.id, date=d, value="Set"))
    s.commit()
    return redirect(url_for("list_employees"))


@app.route("/admin/employees")
def list_employees():
    with SessionLocal() as s:
        employees = {}
        sections = list(s.scalars(select(Section)))
        for sec in sections:
            employees[sec.name] = list(s.scalars(select(Employee).where(Employee.section_id == sec.id)))
    shift_options = {
        "Breakfast Bar": [o for o in BREAKFAST_SHIFTS],
        "Front Desk": [o for o in FRONT_DESK_SHIFTS],
        "Shuttle": [o for o in SHUTTLE_SHIFTS],
    }
    return render_template("employees.html", employees=employees, roles=sections, shift_options=shift_options)


def employee_roles_for(s: Session, eid: int) -> list[Section]:
    role_section_ids = [row[0] for row in s.execute(select(EmployeeRole.section_id).where(EmployeeRole.employee_id == eid)).all()]
    return list(s.scalars(select(Section).where(Section.id.in_(role_section_ids)))) if role_section_ids else []


@app.route("/admin/employees/<int:eid>/roles", methods=["GET", "POST"]) 
def manage_employee_roles(eid: int):
    with SessionLocal() as s:
        emp = s.get(Employee, eid)
        if not emp:
            return redirect(url_for('list_employees'))
        all_sections = list(s.scalars(select(Section)))
        primary_section_id = emp.section_id
        current_secondary = {er.section_id for er in s.scalars(select(EmployeeRole).where(EmployeeRole.employee_id == eid))}

        if request.method == "POST":
            # Parse selected secondary roles
            selected = set(int(x) for x in request.form.getlist("secondary_roles"))
            # Remove primary if accidentally included
            if primary_section_id in selected:
                selected.discard(primary_section_id)
            # Replace all rows to avoid complex diffing errors
            s.execute(delete(EmployeeRole).where(EmployeeRole.employee_id == eid))
            for sid in selected:
                s.add(EmployeeRole(employee_id=eid, section_id=sid))
            s.commit()
            return redirect(url_for('list_employees'))

        return render_template(
            "employee_roles.html",
            employee=emp,
            sections=all_sections,
            primary_id=primary_section_id,
            selected_ids=current_secondary,
        )


@app.route("/admin/employees/<int:eid>/role", methods=["POST"])
def change_employee_role(eid: int):
    role = (request.form.get("role") or "").strip()
    if not role:
        return redirect(url_for('list_employees'))
    with SessionLocal() as s:
        emp = s.get(Employee, eid)
        if not emp:
            return redirect(url_for('list_employees'))
        sec = s.scalar(select(Section).where(Section.name == role))
        if not sec:
            return redirect(url_for('list_employees'))
        # Update role
        emp.section_id = sec.id
        # Reset current week assignments to Set to avoid shift-type mismatch
        week = s.scalar(select(Week).where(Week.start_date == date(2025, 9, 18)))
        for d in daterange(week.start_date, 7):
            a = s.scalar(select(Assignment).where(Assignment.week_id == week.id, Assignment.employee_id == emp.id, Assignment.date == d))
            if a:
                a.value = "Set"
        s.commit()
        # Re-apply time off
        sync_timeoff_to_assignments(week.id, s)
    return redirect(url_for('list_employees'))


@app.route("/admin/employees/<int:eid>/update", methods=["POST"])
def update_employee(eid: int):
    role = (request.form.get("role") or "").strip()
    availability = (request.form.get("availability") or "").strip()
    preferred_shift = (request.form.get("preferred_shift") or "").strip() or None
    seniority_raw = (request.form.get("seniority") or "").strip()
    seniority = int(seniority_raw) if seniority_raw.isdigit() else None
    with SessionLocal() as s:
        emp = s.get(Employee, eid)
        if not emp:
            return redirect(url_for('list_employees'))
        # Role change
        if role:
            sec = s.scalar(select(Section).where(Section.name == role))
            if sec and sec.id != emp.section_id:
                emp.section_id = sec.id
                # Reset current week assignments to Set due to role change
                week = s.scalar(select(Week).where(Week.start_date == date(2025, 9, 18)))
                for d in daterange(week.start_date, 7):
                    a = s.scalar(select(Assignment).where(Assignment.week_id == week.id, Assignment.employee_id == emp.id, Assignment.date == d))
                    if a:
                        a.value = "Set"
        emp.availability = availability or None
        emp.preferred_shift = preferred_shift
        emp.seniority = seniority
        s.commit()
        # Re-apply time off after any changes
        week = s.scalar(select(Week).where(Week.start_date == date(2025, 9, 18)))
        sync_timeoff_to_assignments(week.id, s)
    return redirect(url_for('list_employees'))

@app.route("/admin/employees/<int:eid>/delete", methods=["POST"])
def delete_employee(eid: int):
    with SessionLocal() as s:
        emp = s.get(Employee, eid)
        if not emp:
            return redirect(url_for('list_employees'))
        # Delete assignments for this employee
        s.query(Assignment).filter(Assignment.employee_id == emp.id).delete()
        # Delete availability
        s.query(EmployeeAvailability).filter(EmployeeAvailability.employee_id == emp.id).delete()
        # Delete time off entries matching this name (name-based linkage)
        s.query(TimeOff).filter(TimeOff.name == emp.name).delete()
        # Delete employee
        s.delete(emp)
        s.commit()
    return redirect(url_for('list_employees'))


def role_shift_variants(role_name: str) -> list[str]:
    if role_name == "Breakfast Bar":
        return [s for s in BREAKFAST_SHIFTS if s not in ("Set", TIME_OFF_LABEL)]
    if role_name == "Front Desk":
        return [s for s in FRONT_DESK_SHIFTS if s not in ("Set", TIME_OFF_LABEL)]
    if role_name == "Shuttle":
        return [s for s in SHUTTLE_SHIFTS if s not in ("Set", TIME_OFF_LABEL)]
    return []


def role_availability_variants(role_name: str) -> list[str]:
    # Compact set for availability UI; Front Desk collapses to base variants
    if role_name == "Breakfast Bar":
        return ["5AM–12PM", "6AM–12PM", "7AM–12PM"]
    if role_name == "Front Desk":
        return ["AM", "PM", "Audit"]
    if role_name == "Shuttle":
        return ["AM (3:30AM–11:30AM)", "Midday (10:30AM–6:30PM)", "PM (5:30PM–1:30AM)", "Crew (5:45PM–1:45AM)"]
    return []


@app.route("/admin/employees/<int:eid>/availability", methods=["GET", "POST"]) 
def employee_availability(eid: int):
    days = [
        {"idx": 0, "label": "Mon"},
        {"idx": 1, "label": "Tue"},
        {"idx": 2, "label": "Wed"},
        {"idx": 3, "label": "Thu"},
        {"idx": 4, "label": "Fri"},
        {"idx": 5, "label": "Sat"},
        {"idx": 6, "label": "Sun"},
    ]
    with SessionLocal() as s:
        emp = s.get(Employee, eid)
        if not emp:
            return redirect(url_for('list_employees'))
        # Allow switching role context via query param for multi-role availability
        allowed_sections = [s.get(Section, emp.section_id)] + employee_roles_for(s, eid)
        allowed_names = [sec.name for sec in allowed_sections if sec]
        role = request.args.get("role") or (allowed_sections[0].name if allowed_sections else "")
        if role not in allowed_names and allowed_sections:
            role = allowed_sections[0].name
        variants = role_availability_variants(role)

        if request.method == "POST":
            # Save seniority and preferred shift
            seniority_raw = (request.form.get("seniority") or "").strip()
            emp.seniority = int(seniority_raw) if seniority_raw.isdigit() else None
            pref = (request.form.get("preferred_shift") or "").strip() or None
            # Only allow preferred shift that belongs to availability variants
            emp.preferred_shift = pref if (pref in variants) else None
            # Save weekly shift preferences
            pref_count_raw = (request.form.get("preferred_shifts_per_week") or "").strip()
            max_count_raw = (request.form.get("max_shifts_per_week") or "").strip()
            emp.preferred_shifts_per_week = int(pref_count_raw) if pref_count_raw.isdigit() else None
            emp.max_shifts_per_week = int(max_count_raw) if max_count_raw.isdigit() else None
            # Save availability: clear and re-add
            # NOTE: We don't clear all availability across roles; only rebuild for the role being edited.
            s.query(EmployeeAvailability).filter(
                EmployeeAvailability.employee_id == emp.id,
                EmployeeAvailability.shift_label.in_(variants)
            ).delete()
            selected = request.form.getlist("avail")  # values: "day::label"
            for token in selected:
                try:
                    d_str, label = token.split("::", 1)
                    d_idx = int(d_str)
                except ValueError:
                    continue
                if label in variants and 0 <= d_idx <= 6:
                    s.add(EmployeeAvailability(employee_id=emp.id, day_of_week=d_idx, shift_label=label, allowed=True))
            s.commit()
            return redirect(url_for('employee_availability', eid=eid, role=role))

        # GET: build checked set
        existing = {(ea.day_of_week, ea.shift_label) for ea in s.scalars(select(EmployeeAvailability).where(EmployeeAvailability.employee_id == emp.id))}
        # Render form
        return render_template(
            "employee_availability.html",
            employee=emp,
            role=role,
            variants=variants,
            days=days,
            existing=existing,
            allowed_roles=allowed_names,
        )


@app.route("/timeoff")
def timeoff_page():
    with SessionLocal() as s:
        items = []
        for t in s.scalars(select(TimeOff)):
            same_day = t.from_date == t.to_date
            label = t.from_date.strftime("%b %d") if same_day else f"{t.from_date.strftime('%b %d')} to {t.to_date.strftime('%b %d')}"
            items.append({"id": t.id, "name": t.name, "role": t.role, "label": label, "approved": t.approved, "vacation": bool(getattr(t, "vacation", False))})
        employees = list(s.scalars(select(Employee)))
    return render_template("timeoff.html", time_off=items, employees=employees)


@app.route("/timeoff/new", methods=["GET", "POST"])
def timeoff_new():
    if request.method == 'GET':
        with SessionLocal() as s:
            employees = list(s.scalars(select(Employee)))
        return render_template("timeoff_new.html", employees=employees)
    # POST
    name = (request.form.get('name') or '').strip()
    from_s = (request.form.get('from') or '').strip()
    to_s = (request.form.get('to') or '').strip()
    approved = bool(request.form.get('approved'))
    vacation = bool(request.form.get('vacation'))
    if not name or not from_s or not to_s:
        with SessionLocal() as s:
            employees = list(s.scalars(select(Employee)))
        return render_template("timeoff_new.html", employees=employees, error="All fields are required"), 400
    try:
        fy, fm, fd = map(int, from_s.split('-'))
        ty, tm, td = map(int, to_s.split('-'))
        from_d, to_d = date(fy, fm, fd), date(ty, tm, td)
    except Exception:
        with SessionLocal() as s:
            employees = list(s.scalars(select(Employee)))
        return render_template("timeoff_new.html", employees=employees, error="Invalid dates"), 400
    if to_d < from_d:
        with SessionLocal() as s:
            employees = list(s.scalars(select(Employee)))
        return render_template("timeoff_new.html", employees=employees, error="End date must be after start date"), 400
    with SessionLocal() as s:
        emp = s.scalar(select(Employee).where(Employee.name == name))
        role = s.get(Section, emp.section_id).name if emp else 'Unknown'
        s.add(TimeOff(name=name, role=role, from_date=from_d, to_date=to_d, approved=approved, vacation=vacation))
        s.commit()
        # If approved, update assignments for current week
        if approved:
            week = s.scalar(select(Week).where(Week.start_date == date(2025, 9, 18)))
            if emp and week:
                for d in daterange(week.start_date, 7):
                    if from_d <= d <= to_d:
                        a = s.scalar(select(Assignment).where(Assignment.week_id == week.id, Assignment.employee_id == emp.id, Assignment.date == d))
                        if a:
                            a.value = TIME_OFF_LABEL
                s.commit()
    return redirect(url_for('timeoff_page'))


@app.route("/schedules")
def schedules_page():
    with SessionLocal() as s:
        # Group weeks into 4-week periods anchored to the baseline Thursday
        baseline = date(2025, 9, 18)
        periods: dict[int, dict] = {}
        for w in s.scalars(select(Week)):
            # Compute 4-week period index relative to baseline (28 days)
            delta_days = (w.start_date - baseline).days
            idx = delta_days // 28 if delta_days >= 0 else -((-delta_days - 1) // 28) - 1
            if idx not in periods:
                period_start = baseline + timedelta(days=idx * 28)
                period_end = period_start + timedelta(days=27)
                periods[idx] = {
                    "first_week_id": w.id,
                    "start": period_start,
                    "end": period_end,
                }
            else:
                # Ensure first_week_id points to the earliest week in the period
                existing_week = s.get(Week, periods[idx]["first_week_id"]) if periods[idx].get("first_week_id") else None
                if existing_week and w.start_date < existing_week.start_date:
                    periods[idx]["first_week_id"] = w.id
        # Build list sorted by start date descending (most recent first)
        items = []
        for p in sorted(periods.values(), key=lambda x: x["start"], reverse=True):
            items.append({
                "id": p["first_week_id"],
                "start": p["start"].strftime('%b %d, %Y'),
                "end": p["end"].strftime('%b %d, %Y'),
            })
    return render_template("schedules.html", periods=items)

@app.route("/assign", methods=["POST"])
def assign():
    data = request.get_json(force=True)
    section = data.get("section")
    employee_name = data.get("employee")
    date_key = data.get("date")
    value = data.get("value")
    week_id = data.get("week_id")  # Get week_id from request

    with SessionLocal() as s:
        # Use provided week_id or fall back to default week
        if week_id:
            week = s.get(Week, week_id)
        else:
            week = s.scalar(select(Week).where(Week.start_date == date(2025, 9, 18)))
        
        if not week:
            return jsonify({"ok": False, "error": "Week not found"}), 404
            
        # Resolve employee and allow secondary role assignments
        sec = s.scalar(select(Section).where(Section.name == section))
        if not sec:
            return jsonify({"ok": False, "error": "Unknown section"}), 400
        emp = s.scalar(select(Employee).where(Employee.name == employee_name))
        if not emp:
            return jsonify({"ok": False, "error": "Unknown employee"}), 400
        # Check permission: primary in section OR has secondary role mapping
        allowed = emp.section_id == sec.id or (s.scalar(select(EmployeeRole).where(EmployeeRole.employee_id == emp.id, EmployeeRole.section_id == sec.id).limit(1)) is not None)
        if not allowed:
            return jsonify({"ok": False, "error": "Employee not allowed for this section"}), 400
        try:
            y, m, d = map(int, date_key.split("-"))
            dte = date(y, m, d)
        except Exception:
            return jsonify({"ok": False, "error": "Bad date"}), 400

        a = s.scalar(select(Assignment).where(Assignment.week_id == week.id, Assignment.employee_id == emp.id, Assignment.date == dte))
        if not a:
            return jsonify({"ok": False, "error": "Assignment not found"}), 404
        # Block assignment if approved time off; force TIME OFF
        if has_approved_timeoff(emp.name, dte, s):
            if a.value != TIME_OFF_LABEL:
                a.value = TIME_OFF_LABEL
                s.commit()
            counts, missing, required, variant_counts, shuttle_missing, shuttle_required, shuttle_counts, bb_missing, bb_required, bb_counts, fd_duplicates = coverage_snapshot_db(week.id)
            return jsonify({
                "ok": False,
                "error": "Approved time off",
                "code": "timeoff",
                "value": TIME_OFF_LABEL,
                "counts": counts,
                "missing": missing,
                "required": required,
                "variant_counts": variant_counts,
                "shuttle_missing": shuttle_missing,
                "shuttle_required": shuttle_required,
                "shuttle_counts": shuttle_counts,
                "bb_missing": bb_missing,
                "bb_required": bb_required,
                "bb_counts": bb_counts,
                "fd_duplicates": fd_duplicates,
                "double_booked": double_booked_snapshot(week.id),
            }), 409
        a.value = value
        s.commit()

    counts, missing, required, variant_counts, shuttle_missing, shuttle_required, shuttle_counts, bb_missing, bb_required, bb_counts, fd_duplicates = coverage_snapshot_db(week.id)
    return jsonify({
        "ok": True,
        "counts": counts,
        "missing": missing,
        "required": required,
        "variant_counts": variant_counts,
        "shuttle_missing": shuttle_missing,
        "shuttle_required": shuttle_required,
        "shuttle_counts": shuttle_counts,
        "bb_missing": bb_missing,
        "bb_required": bb_required,
        "bb_counts": bb_counts,
        "fd_duplicates": fd_duplicates,
        "double_booked": double_booked_snapshot(week.id),
    })


@app.route("/timeoff/delete/<int:tid>", methods=["POST"])
def delete_timeoff(tid):
    with SessionLocal() as s:
        to = s.get(TimeOff, tid)
        if not to:
            return jsonify({"ok": False, "error": "Not found"}), 404
        
        # Get the week reference for coverage updates
        week = s.scalar(select(Week).where(Week.start_date == date(2025, 9, 18)))
        
        # If it was approved, we need to update assignments back to "Set"
        if to.approved and week:
            emp = s.scalar(select(Employee).where(Employee.name == to.name))
            if emp:
                for d in daterange(week.start_date, 7):
                    in_range = to.from_date <= d <= to.to_date
                    if in_range:
                        a = s.scalar(select(Assignment).where(Assignment.week_id == week.id, Assignment.employee_id == emp.id, Assignment.date == d))
                        if a and a.value == TIME_OFF_LABEL:
                            a.value = "Set"
        
        s.delete(to)
        s.commit()
        
        # Only update coverage if we have a valid week
        if week:
            counts, missing, required, variant_counts, shuttle_missing, shuttle_required, shuttle_counts, bb_missing, bb_required, bb_counts, fd_duplicates = coverage_snapshot_db(week.id)
            return jsonify({
                "ok": True,
                "counts": counts,
                "missing": missing,
                "required": required,
                "variant_counts": variant_counts,
                "shuttle_missing": shuttle_missing,
                "shuttle_required": shuttle_required,
                "shuttle_counts": shuttle_counts,
                "bb_missing": bb_missing,
                "bb_required": bb_required,
                "bb_counts": bb_counts,
                "fd_duplicates": fd_duplicates,
                "double_booked": double_booked_snapshot(week.id),
            })
        else:
            return jsonify({"ok": True, "counts": {}, "missing": {}, "required": 0, "variant_counts": {}})


@app.route("/timeoff/toggle", methods=["POST"])
def toggle_timeoff():
    data = request.get_json(force=True)
    tid = data.get("id")
    approved = bool(data.get("approved"))
    with SessionLocal() as s:
        to = s.get(TimeOff, tid)
        if not to:
            return jsonify({"ok": False, "error": "Not found"}), 404
        to.approved = approved
        # Update assignments for the current primary week only
        week = s.scalar(select(Week).where(Week.start_date == date(2025, 9, 18)))
        emp = s.scalar(select(Employee).where(Employee.name == to.name))
        if emp:
            for d in daterange(week.start_date, 7):
                in_range = to.from_date <= d <= to.to_date
                a = s.scalar(select(Assignment).where(Assignment.week_id == week.id, Assignment.employee_id == emp.id, Assignment.date == d))
                if not a:
                    continue
                if approved and in_range:
                    a.value = TIME_OFF_LABEL
                elif not approved and a.value == TIME_OFF_LABEL and in_range:
                    a.value = "Set"
        s.commit()
        counts, missing, required, variant_counts, shuttle_missing, shuttle_required, shuttle_counts, bb_missing, bb_required, bb_counts = coverage_snapshot_db(week.id)
        item = {
            "id": to.id,
            "name": to.name,
            "role": to.role,
            "from": to.from_date.isoformat(),
            "to": to.to_date.isoformat(),
            "approved": to.approved,
        }
    return jsonify({
        "ok": True,
        "item": item,
        "counts": counts,
        "missing": missing,
        "required": required,
        "variant_counts": variant_counts,
        "shuttle_missing": shuttle_missing,
        "shuttle_required": shuttle_required,
        "shuttle_counts": shuttle_counts,
        "bb_missing": bb_missing,
        "bb_required": bb_required,
        "bb_counts": bb_counts,
        "fd_duplicates": fd_duplicates,
        "double_booked": double_booked_snapshot(week.id),
    })


@app.route("/timeoff/vacation", methods=["POST"])
def toggle_timeoff_vacation():
    data = request.get_json(force=True)
    tid = data.get("id")
    vacation = bool(data.get("vacation"))
    with SessionLocal() as s:
        to = s.get(TimeOff, tid)
        if not to:
            return jsonify({"ok": False, "error": "Not found"}), 404
        to.vacation = vacation
        s.commit()
        return jsonify({
            "ok": True,
            "item": {
                "id": to.id,
                "name": to.name,
                "role": to.role,
                "from": to.from_date.isoformat(),
                "to": to.to_date.isoformat(),
                "approved": bool(to.approved),
                "vacation": bool(to.vacation),
            }
        })


def generate_new_schedule_db(week_id: int):
    """Legacy: generate a single week (kept for reference)."""
    with SessionLocal() as s:
        fd_emp_ids = list(s.scalars(select(Employee.id).join(Section).where(Section.name == "Front Desk")))
        bb_emp_ids = list(s.scalars(select(Employee.id).join(Section).where(Section.name == "Breakfast Bar")))
        sh_emp_ids = list(s.scalars(select(Employee.id).join(Section).where(Section.name == "Shuttle")))
        wk = s.get(Week, week_id)

        # Reset all
        s.execute(update(Assignment).where(Assignment.week_id == week_id).values(value="Set"))
        s.commit()

        # Breakfast Bar: ensure one person at 5am, 6am, and 7am each day
        for d in daterange(wk.start_date, 7):
            pool = bb_emp_ids[:]
            shifts = ["5AM–12PM", "6AM–12PM", "7AM–12PM"]
            
            # Assign one person to each of the three morning shifts
            for shift in shifts:
                if pool:  # Only assign if there are available employees
                    # pick someone without time off request
                    eligible = [eid for eid in pool if not has_any_timeoff(s.get(Employee, eid).name if s.get(Employee, eid) else '', d, s)]
                    if not eligible:
                        continue
                    eid = sample(eligible, k=1)[0]
                    pool.remove(eid)
                    a = s.scalar(select(Assignment).where(Assignment.week_id == week_id, Assignment.employee_id == eid, Assignment.date == d))
                    if a:
                        a.value = shift
            
            # Remaining employees get "Set" (no shift)
            for eid in pool:
                a = s.scalar(select(Assignment).where(Assignment.week_id == week_id, Assignment.employee_id == eid, Assignment.date == d))
                if a:
                    a.value = "Set"

        # Front Desk: ensure 2 per variant (AM/PM/Audit) per day
        for d in daterange(wk.start_date, 7):
            pool = fd_emp_ids[:]
            # AM
            am_eligible = [eid for eid in pool if not has_any_timeoff(s.get(Employee, eid).name if s.get(Employee, eid) else '', d, s)]
            am = sample(am_eligible, k=min(2, len(am_eligible))) if am_eligible else []
            # Senior earlier: sort by known seniority order
            def _sen_rank_fd(eid: int) -> int:
                name = s.get(Employee, eid).name if s.get(Employee, eid) else ''
                # Exception: Ryan should be treated as least-preferred for earlier start
                if name == 'Ryan':
                    return 10000
                try:
                    return SENIORITY_ORDER.index(name)
                except ValueError:
                    return 9999
            am.sort(key=_sen_rank_fd)
            for i, eid in enumerate(am):
                pool.remove(eid)
                a = s.scalar(select(Assignment).where(Assignment.week_id == week_id, Assignment.employee_id == eid, Assignment.date == d))
                if a:
                    # If only Ryan was picked, give him the later stagger
                    if len(am) == 1 and (s.get(Employee, eid).name if s.get(Employee, eid) else '') == 'Ryan':
                        a.value = "AM (6:15AM–2:15PM)"
                    else:
                        a.value = "AM (6:00AM–2:00PM)" if i == 0 else "AM (6:15AM–2:15PM)"
            # PM
            pm_candidates = [eid for eid in pool if not has_any_timeoff(s.get(Employee, eid).name if s.get(Employee, eid) else '', d, s)]
            pm = sample(pm_candidates, k=min(2, len(pm_candidates))) if pm_candidates else []
            pm.sort(key=_sen_rank_fd)
            for i, eid in enumerate(pm):
                pool.remove(eid)
                a = s.scalar(select(Assignment).where(Assignment.week_id == week_id, Assignment.employee_id == eid, Assignment.date == d))
                if a:
                    if len(pm) == 1 and (s.get(Employee, eid).name if s.get(Employee, eid) else '') == 'Ryan':
                        a.value = "PM (2:15PM–10:15PM)"
                    else:
                        a.value = "PM (2:00PM–10:00PM)" if i == 0 else "PM (2:15PM–10:15PM)"
            # Audit
            au_candidates = [eid for eid in pool if not has_any_timeoff(s.get(Employee, eid).name if s.get(Employee, eid) else '', d, s)]
            au = sample(au_candidates, k=min(2, len(au_candidates))) if au_candidates else []
            au.sort(key=_sen_rank_fd)
            for i, eid in enumerate(au):
                pool.remove(eid)
                a = s.scalar(select(Assignment).where(Assignment.week_id == week_id, Assignment.employee_id == eid, Assignment.date == d))
                if a:
                    if len(au) == 1 and (s.get(Employee, eid).name if s.get(Employee, eid) else '') == 'Ryan':
                        a.value = "Audit (10:15PM–6:15AM)"
                    else:
                        a.value = "Audit (10:00PM–6:00AM)" if i == 0 else "Audit (10:15PM–6:15AM)"

        # Shuttle: one agent per variant per day (as available)
        for d in daterange(wk.start_date, 7):
            pool = sh_emp_ids[:]
            variants = [
                "AM (3:30AM–11:30AM)",
                "Midday (10:30AM–6:30PM)",
                "PM (5:30PM–1:30AM)",
                "Crew (5:45PM–1:45AM)",
            ]
            for v in variants:
                if not pool:
                    break
                eligible = [eid for eid in pool if not has_any_timeoff(s.get(Employee, eid).name if s.get(Employee, eid) else '', d, s)]
                if not eligible:
                    continue
                eid = sample(eligible, k=1)[0]
                pool.remove(eid)
                a = s.scalar(select(Assignment).where(Assignment.week_id == week_id, Assignment.employee_id == eid, Assignment.date == d))
                if a:
                    a.value = v

        # Ensure approved time off overrides any generated shifts
        sync_timeoff_to_assignments(week_id, s)

        s.commit()


def _ensure_week_and_assignments(s: Session, start_d: date) -> Week:
    wk = s.scalar(select(Week).where(Week.start_date == start_d))
    if not wk:
        wk = Week(start_date=start_d)
        s.add(wk)
        s.flush()
    # Ensure assignments exist for all employees/days
    for emp in s.scalars(select(Employee)):
        for d in daterange(start_d, 7):
            exists = s.scalar(select(Assignment).where(Assignment.week_id == wk.id, Assignment.employee_id == emp.id, Assignment.date == d))
            if not exists:
                s.add(Assignment(week_id=wk.id, employee_id=emp.id, date=d, value="Set"))
    s.commit()
    return wk


def _build_availability_index(s: Session) -> dict[int, set[tuple[int, str]]]:
    idx: dict[int, set[tuple[int, str]]] = {}
    for ea in s.scalars(select(EmployeeAvailability)):
        idx.setdefault(ea.employee_id, set()).add((ea.day_of_week, ea.shift_label))
    return idx


def _fd_variant(label: str) -> str:
    if label.startswith("AM"):
        return "AM"
    if label.startswith("PM"):
        return "PM"
    if label.startswith("Audit"):
        return "Audit"
    return label


def _is_available(emp_id: int, role: str, shift_label: str, dte: date, avail_idx: dict[int, set[tuple[int, str]]]) -> bool:
    allowed = avail_idx.get(emp_id)
    # If no availability configured, treat as unavailable
    # This prevents scheduling employees who have not set any availability.
    if not allowed:
        return False
    dow = dte.weekday()  # 0=Mon
    if role == "Front Desk":
        key = (dow, _fd_variant(shift_label))
    else:
        key = (dow, shift_label)
    return key in allowed


def _pref_token(role: str, label: str) -> str:
    if role == "Front Desk":
        return _fd_variant(label)
    return label


def _effective_max_for_role(emp_info: dict[int, dict], eid: int, role: str) -> Optional[int]:
    """Return the max shifts/week for this employee in the context of a role.
    If employee-specific max is set, use it; otherwise default per role:
    - Front Desk: 5
    - Shuttle: 4
    - Breakfast Bar: no default cap (None)
    """
    maxw = emp_info[eid].get("max_per_week")
    if maxw is not None:
        return maxw
    if role == "Front Desk":
        return 5
    if role == "Shuttle":
        return 4
    return None


def generate_4_week_schedule(start_week_id: int):
    """Generate a 4-week schedule starting at the given week.
    Respects EmployeeAvailability; leaves unfilled if no available employees.
    Also reapplies approved time off per week.
    """
    with SessionLocal() as s:
        base_week = s.get(Week, start_week_id)
        if not base_week:
            return

        # Simple file logger to capture scheduling decisions
        def _log(msg: str):
            try:
                with open('log.txt', 'a', encoding='utf-8') as _f:
                    _f.write(msg.rstrip() + "\n")
            except Exception:
                pass

        # Truncate previous log
        try:
            open('log.txt', 'w').close()
        except Exception:
            pass
        _log(f"Generate 4-week schedule starting week_id={start_week_id} ({base_week.start_date.isoformat()})")

        # Precompute availability and role mappings
        avail_idx = _build_availability_index(s)
        emp_info: dict[int, dict] = {}
        emp_name: dict[int, str] = {}
        for e in s.scalars(select(Employee)):
            sec = s.get(Section, e.section_id)
            emp_info[e.id] = {
                "role": (sec.name if sec else ""),
                "seniority": e.seniority or 0,
                "preferred_shift": e.preferred_shift or None,
                "pref_per_week": e.preferred_shifts_per_week if (e.preferred_shifts_per_week is not None) else None,
                "max_per_week": e.max_shifts_per_week if (e.max_shifts_per_week is not None) else None,
            }
            emp_name[e.id] = e.name

        # Gather employees by section including secondary roles
        def employees_for_section(name: str) -> list[int]:
            sec = s.scalar(select(Section).where(Section.name == name))
            if not sec:
                return []
            primary_ids = list(s.scalars(select(Employee.id).where(Employee.section_id == sec.id)))
            secondary_ids = [row[0] for row in s.execute(select(EmployeeRole.employee_id).where(EmployeeRole.section_id == sec.id)).all()]
            # Combine and dedupe
            return list({*(primary_ids), *secondary_ids})

        fd_emp_ids = employees_for_section("Front Desk")
        bb_emp_ids = employees_for_section("Breakfast Bar")
        sh_emp_ids = employees_for_section("Shuttle")
        _log(f"FD-capable: {[emp_name.get(i,'?') for i in fd_emp_ids]}")
        _log(f"BB-capable: {[emp_name.get(i,'?') for i in bb_emp_ids]}")
        _log(f"SH-capable: {[emp_name.get(i,'?') for i in sh_emp_ids]}")

        for week_offset in range(4):
            start_d = base_week.start_date + timedelta(days=7 * week_offset)
            wk = _ensure_week_and_assignments(s, start_d)

            # Reset all to Set and clear dismissed flags (if column exists)
            if hasattr(Assignment, 'dismissed_timeoff'):
                s.execute(update(Assignment).where(Assignment.week_id == wk.id).values(value="Set", dismissed_timeoff=0))
            else:
                s.execute(update(Assignment).where(Assignment.week_id == wk.id).values(value="Set"))
            s.commit()

            # Track assigned count per employee within this week (excludes TIME OFF)
            assigned_counts: dict[int, int] = {eid: 0 for eid in emp_info.keys()}

            # Track last assigned token per employee per role to prefer consistency within the week
            last_token_per_role: dict[tuple[int, str], Optional[str]] = {}

            def rest_ok(eid: int, dte: date, role: str, label: str) -> bool:
                # Apply simple rest constraints across consecutive days, even over week boundaries
                prev_date = dte - timedelta(days=1)
                a_prev = s.scalar(select(Assignment).where(Assignment.employee_id == eid, Assignment.date == prev_date))
                if not a_prev or not a_prev.value or a_prev.value in ("Set", TIME_OFF_LABEL):
                    return True
                # No AM after a previous-day PM (short 8h rest)
                if role == "Front Desk" and label.startswith("AM") and a_prev.value.startswith("PM"):
                    return False
                # No PM after a previous-night Audit (8h rest window)
                if role == "Front Desk" and label.startswith("PM") and a_prev.value.startswith("Audit"):
                    return False
                # Shuttle rest rules
                if role == "Shuttle":
                    # No AM after a previous-day late shift (PM or Crew ends after midnight)
                    if label.startswith("AM") and (a_prev.value.startswith("PM") or a_prev.value.startswith("Crew")):
                        return False
                    # Be cautious with Midday after Crew (only ~8h45m rest)
                    if label.startswith("Midday") and a_prev.value.startswith("Crew"):
                        return False
                return True

            def pick_best(candidates: list[int], role: str, label: str) -> Optional[int]:
                # Filter out those exceeding max per week
                filtered = []
                for eid in candidates:
                    maxw = _effective_max_for_role(emp_info, eid, role)
                    if maxw is not None and assigned_counts[eid] >= maxw:
                        _log(f"{current_loop_date} {role} {label}: drop {emp_name.get(eid,'?')} (max/week reached {maxw})")
                        continue
                    # Rest constraints
                    # Determine actual date in outer calling context by closure hack: rely on current_loop_date variable
                    if not rest_ok(eid, current_loop_date, role, label):
                        _log(f"{current_loop_date} {role} {label}: drop {emp_name.get(eid,'?')} (rest rule)")
                        continue
                    filtered.append(eid)
                if not filtered:
                    _log(f"{current_loop_date} {role} {label}: no candidates after filters")
                    return None
                def key_fn(eid: int):
                    under_pref = False
                    prefw = emp_info[eid]["pref_per_week"]
                    if prefw is not None:
                        under_pref = assigned_counts[eid] < prefw
                    # Compare this employee's own preferred shift against the current token
                    want = emp_info[eid]["preferred_shift"]
                    preferred_match = (want is not None and _pref_token(role, label) == want)
                    # Prefer consistency with last assigned token within the week for this role
                    last_tok = last_token_per_role.get((eid, role))
                    consistent = (last_tok is not None and last_tok == _pref_token(role, label))
                    return (
                        0 if under_pref else 1,
                        0 if preferred_match else 1,
                        0 if consistent else 1,
                        -(emp_info[eid]["seniority"] or 0),
                        assigned_counts[eid],
                    )
                ranked = sorted(filtered, key=key_fn)
                _log(f"{current_loop_date} {role} {label}: ranked {[emp_name.get(e,'?') for e in ranked]}")
                best = ranked[0]
                return best

            # Helper to mark if candidate dropped solely due to time off
            def _mark_dismissed(eid: int, dte: date):
                if not hasattr(Assignment, 'dismissed_timeoff'):
                    return
                a = s.scalar(select(Assignment).where(Assignment.week_id == wk.id, Assignment.employee_id == eid, Assignment.date == dte))
                if a and hasattr(a, 'dismissed_timeoff'):
                    a.dismissed_timeoff = 1

            # Front Desk: 2 per AM/PM/Audit per day if available (prioritize FD first)
            for d in daterange(wk.start_date, 7):
                current_loop_date = d
                pool = fd_emp_ids[:]
                for variant, opts in (
                    ("AM", ["AM (6:00AM–2:00PM)", "AM (6:15AM–2:15PM)"]),
                    ("PM", ["PM (2:00PM–10:00PM)", "PM (2:15PM–10:15PM)"]),
                    ("Audit", ["Audit (10:00PM–6:00AM)", "Audit (10:15PM–6:15AM)"]),
                ):
                    # Choose up to 2 employees available and not already assigned that day
                    cand = []
                    _log(f"{d} FD {variant}: start pool {[emp_name.get(i,'?') for i in pool]}")
                    for eid in pool:
                        a = s.scalar(select(Assignment).where(Assignment.week_id == wk.id, Assignment.employee_id == eid, Assignment.date == d))
                        if not a or a.value != "Set":
                            _log(f"{d} FD {variant}: drop {emp_name.get(eid,'?')} (already assigned {a.value if a else 'None'})")
                            continue
                        if not _is_available(eid, "Front Desk", variant, d, avail_idx):
                            _log(f"{d} FD {variant}: drop {emp_name.get(eid,'?')} (not available)")
                            continue
                        if not rest_ok(eid, d, "Front Desk", variant):
                            _log(f"{d} FD {variant}: drop {emp_name.get(eid,'?')} (rest rule)")
                            continue
                        # Skip anyone with any time off request on this date
                        if has_any_timeoff(emp_name.get(eid, ''), d, s):
                            _log(f"{d} FD {variant}: drop {emp_name.get(eid,'?')} (time off request)")
                            _mark_dismissed(eid, d)
                            continue
                        if _effective_max_for_role(emp_info, eid, "Front Desk") is not None and assigned_counts[eid] >= _effective_max_for_role(emp_info, eid, "Front Desk"):
                            _log(f"{d} FD {variant}: drop {emp_name.get(eid,'?')} (max/week reached)")
                            continue
                        # candidate accepted
                        cand.append(eid)
                    _log(f"{d} FD {variant}: candidates {[emp_name.get(i,'?') for i in cand]}")
                    if not cand:
                        continue
                    picks: list[int] = []
                    for slot_i in range(2):
                        pick = pick_best([e for e in cand if e not in picks], "Front Desk", variant)
                        if pick is None:
                            break
                        picks.append(pick)
                    # Senior earlier: ensure earlier stagger goes to more senior
                    def _seniority_rank(eid: int) -> int:
                        name = emp_name.get(eid, '')
                        # Exception: Ryan should get the later stagger, so sort him last
                        if name == 'Ryan':
                            return 10000
                        try:
                            return SENIORITY_ORDER.index(name)
                        except ValueError:
                            return 9999
                    picks.sort(key=_seniority_rank)
                    for i, eid in enumerate(picks):
                        pool.remove(eid)
                        # If only Ryan was picked for this variant, give him the later stagger
                        if len(picks) == 1 and emp_name.get(eid, '') == 'Ryan' and len(opts) > 1:
                            label = opts[1]
                        else:
                            label = opts[i] if i < len(opts) else opts[-1]
                        a = s.scalar(select(Assignment).where(Assignment.week_id == wk.id, Assignment.employee_id == eid, Assignment.date == d))
                        if a:
                            a.value = label
                            assigned_counts[eid] += 1
                            last_token_per_role[(eid, "Front Desk")] = _pref_token("Front Desk", variant)
                            _log(f"{d} FD assign {label} -> {emp_name.get(eid,'?')}")

            # After FD assignments, strictly reserve FD-capable employees when FD coverage is still missing
            reserved_fd_by_date: dict[date, set[int]] = {}
            def fd_missing_on(day: date) -> bool:
                fd_counts = {"AM": 0, "PM": 0, "Audit": 0}
                for eid in fd_emp_ids:
                    a = s.scalar(select(Assignment).where(Assignment.week_id == wk.id, Assignment.employee_id == eid, Assignment.date == day))
                    if not a or not a.value or a.value in ("Set", TIME_OFF_LABEL):
                        continue
                    if a.value.startswith("AM"):
                        fd_counts["AM"] += 1
                    elif a.value.startswith("PM"):
                        fd_counts["PM"] += 1
                    elif a.value.startswith("Audit"):
                        fd_counts["Audit"] += 1
                missing = any(fd_counts[v] < 2 for v in ("AM", "PM", "Audit"))
                if missing:
                    _log(f"{day} FD missing counts={fd_counts}")
                return missing
            for d in daterange(wk.start_date, 7):
                # Count current FD coverage among FD-capable employees
                fd_counts = {"AM": 0, "PM": 0, "Audit": 0}
                for eid in fd_emp_ids:
                    a = s.scalar(select(Assignment).where(Assignment.week_id == wk.id, Assignment.employee_id == eid, Assignment.date == d))
                    if not a or not a.value or a.value in ("Set", TIME_OFF_LABEL):
                        continue
                    if a.value.startswith("AM"):
                        fd_counts["AM"] += 1
                    elif a.value.startswith("PM"):
                        fd_counts["PM"] += 1
                    elif a.value.startswith("Audit"):
                        fd_counts["Audit"] += 1
                if any(fd_counts[v] < 2 for v in ("AM", "PM", "Audit")):
                    # Reserve all FD-capable employees who are still unassigned (Set) on this day
                    reserve: set[int] = set()
                    for eid in fd_emp_ids:
                        a = s.scalar(select(Assignment).where(Assignment.week_id == wk.id, Assignment.employee_id == eid, Assignment.date == d))
                        if a and a.value == "Set":
                            reserve.add(eid)
                    if reserve:
                        reserved_fd_by_date[d] = reserve
                        _log(f"{d} FD reserve {[emp_name.get(i,'?') for i in sorted(reserve)]}")

            # Breakfast Bar: one per 5/6/7am per day if available (after FD; skip reserved FD candidates and prefer non-FD when FD is missing)
            for d in daterange(wk.start_date, 7):
                current_loop_date = d
                pool = bb_emp_ids[:]
                for shift in ["5AM–12PM", "6AM–12PM", "7AM–12PM"]:
                    candidates = []
                    for eid in pool:
                        # Respect FD reservation for this day
                        if d in reserved_fd_by_date and eid in reserved_fd_by_date[d]:
                            _log(f"{d} BB {shift}: skip {emp_name.get(eid,'?')} (reserved for FD)")
                            continue
                        # If FD is missing today, avoid using FD-capable people in BB
                        if fd_missing_on(d) and eid in fd_emp_ids:
                            _log(f"{d} BB {shift}: skip {emp_name.get(eid,'?')} (FD missing; FD-capable)")
                            continue
                        a = s.scalar(select(Assignment).where(Assignment.week_id == wk.id, Assignment.employee_id == eid, Assignment.date == d))
                        if not a or a.value != "Set":
                            _log(f"{d} BB {shift}: drop {emp_name.get(eid,'?')} (already assigned {a.value if a else 'None'})")
                            continue
                        if has_any_timeoff(emp_name.get(eid, ''), d, s):
                            _log(f"{d} BB {shift}: drop {emp_name.get(eid,'?')} (time off request)")
                            # Mark dismissed only if they'd otherwise pass availability/rest for this shift
                            if _is_available(eid, "Breakfast Bar", shift, d, avail_idx) and rest_ok(eid, d, "Breakfast Bar", shift):
                                _mark_dismissed(eid, d)
                            continue
                        if _is_available(eid, "Breakfast Bar", shift, d, avail_idx) and rest_ok(eid, d, "Breakfast Bar", shift):
                            candidates.append(eid)
                        else:
                            _log(f"{d} BB {shift}: drop {emp_name.get(eid,'?')} (not available/rest)")
                    pick = pick_best(candidates, "Breakfast Bar", shift)
                    if pick is None:
                        continue
                    pool.remove(pick)
                    a = s.scalar(select(Assignment).where(Assignment.week_id == wk.id, Assignment.employee_id == pick, Assignment.date == d))
                    if a:
                        a.value = shift
                        assigned_counts[pick] += 1
                        last_token_per_role[(pick, "Breakfast Bar")] = _pref_token("Breakfast Bar", shift)
                        _log(f"{d} BB assign {shift} -> {emp_name.get(pick,'?')}")

            # Shuttle: one per variant per day if available (skip reserved FD candidates and prefer non-FD when FD is missing)
            for d in daterange(wk.start_date, 7):
                current_loop_date = d
                pool = sh_emp_ids[:]
                variants = [
                    "AM (3:30AM–11:30AM)",
                    "Midday (10:30AM–6:30PM)",
                    "PM (5:30PM–1:30AM)",
                    "Crew (5:45PM–1:45AM)",
                ]
                for v in variants:
                    cand = []
                    for eid in pool:
                        # Respect FD reservation for this day
                        if d in reserved_fd_by_date and eid in reserved_fd_by_date[d]:
                            _log(f"{d} SH {v}: skip {emp_name.get(eid,'?')} (reserved for FD)")
                            continue
                        # If FD is missing today, avoid using FD-capable people in Shuttle
                        if fd_missing_on(d) and eid in fd_emp_ids:
                            _log(f"{d} SH {v}: skip {emp_name.get(eid,'?')} (FD missing; FD-capable)")
                            continue
                        a = s.scalar(select(Assignment).where(Assignment.week_id == wk.id, Assignment.employee_id == eid, Assignment.date == d))
                        if not a or a.value != "Set":
                            _log(f"{d} SH {v}: drop {emp_name.get(eid,'?')} (already assigned {a.value if a else 'None'})")
                            continue
                        if has_any_timeoff(emp_name.get(eid, ''), d, s):
                            _log(f"{d} SH {v}: drop {emp_name.get(eid,'?')} (time off request)")
                            if rest_ok(eid, d, "Shuttle", v):
                                _mark_dismissed(eid, d)
                            continue
                        if _is_available(eid, "Shuttle", v, d, avail_idx) and rest_ok(eid, d, "Shuttle", v):
                            cand.append(eid)
                        else:
                            _log(f"{d} SH {v}: drop {emp_name.get(eid,'?')} (not available/rest)")
                    pick = pick_best(cand, "Shuttle", v)
                    if pick is None:
                        continue
                    pool.remove(pick)
                    a = s.scalar(select(Assignment).where(Assignment.week_id == wk.id, Assignment.employee_id == pick, Assignment.date == d))
                    if a:
                        a.value = v
                        assigned_counts[pick] += 1
                        last_token_per_role[(pick, "Shuttle")] = _pref_token("Shuttle", v)
                        _log(f"{d} SH assign {v} -> {emp_name.get(pick,'?')}")

            # Ensure approved time off overrides any generated shifts for this week
            sync_timeoff_to_assignments(wk.id, s)

        s.commit()


@app.route("/generate", methods=["POST", "GET"]) 
def generate():
    with SessionLocal() as s:
        week = s.scalar(select(Week).where(Week.start_date == date(2025, 9, 18)))
        week_id = week.id
    generate_4_week_schedule(week_id)
    return redirect(url_for("index"))


@app.route("/week/<int:week_id>/generate", methods=["POST"]) 
def generate_week(week_id: int):
    generate_4_week_schedule(week_id)
    return redirect(url_for("view_week", week_id=week_id))


@app.route("/delete", methods=["POST"]) 
def delete_schedule():
    with SessionLocal() as s:
        week = s.scalar(select(Week).where(Week.start_date == date(2025, 9, 18)))
        if hasattr(Assignment, 'dismissed_timeoff'):
            s.execute(update(Assignment).where(Assignment.week_id == week.id).values(value="Set", dismissed_timeoff=0))
        else:
            s.execute(update(Assignment).where(Assignment.week_id == week.id).values(value="Set"))
        s.commit()
    return redirect(url_for("index"))


@app.route("/week/<int:week_id>/delete", methods=["POST"]) 
def delete_week_schedule(week_id: int):
    with SessionLocal() as s:
        wk = s.get(Week, week_id)
        if not wk:
            return redirect(url_for("schedules_page"))

        # Protect the baseline week: reset instead of deleting
        baseline = date(2025, 9, 18)
        if wk.start_date == baseline:
            s.execute(update(Assignment).where(Assignment.week_id == week_id).values(value="Set"))
            s.commit()
            return redirect(url_for("view_week", week_id=week_id))

        # Delete assignments for this week, then the week itself
        s.execute(delete(Assignment).where(Assignment.week_id == week_id))
        s.delete(wk)
        s.commit()
    return redirect(url_for("schedules_page"))


@app.route("/week/<int:week_id>/prev")
def prev_week(week_id: int):
    """Navigate to the previous week"""
    with SessionLocal() as s:
        current_week = s.get(Week, week_id)
        if not current_week:
            return redirect(url_for("index"))
        
        # Calculate previous week start date (7 days earlier)
        prev_start_date = current_week.start_date - timedelta(days=7)
        
        # Check if previous week exists, create if not
        prev_week = s.scalar(select(Week).where(Week.start_date == prev_start_date))
        if not prev_week:
            prev_week = Week(start_date=prev_start_date)
            s.add(prev_week)
            s.flush()
            
            # Create assignments for all employees for the new week
            for emp in s.scalars(select(Employee)):
                for d in daterange(prev_start_date, 7):
                    s.add(Assignment(week_id=prev_week.id, employee_id=emp.id, date=d, value="Set"))
            s.commit()
            
            # Sync any approved time off for this week
            sync_timeoff_to_assignments(prev_week.id, s)
        
        return redirect(url_for("view_week", week_id=prev_week.id))


@app.route("/week/<int:week_id>/next")
def next_week(week_id: int):
    """Navigate to the next week"""
    with SessionLocal() as s:
        current_week = s.get(Week, week_id)
        if not current_week:
            return redirect(url_for("index"))
        
        # Calculate next week start date (7 days later)
        next_start_date = current_week.start_date + timedelta(days=7)
        
        # Check if next week exists, create if not
        next_week = s.scalar(select(Week).where(Week.start_date == next_start_date))
        if not next_week:
            next_week = Week(start_date=next_start_date)
            s.add(next_week)
            s.flush()
            
            # Create assignments for all employees for the new week
            for emp in s.scalars(select(Employee)):
                for d in daterange(next_start_date, 7):
                    s.add(Assignment(week_id=next_week.id, employee_id=emp.id, date=d, value="Set"))
            s.commit()
            
            # Sync any approved time off for this week
            sync_timeoff_to_assignments(next_week.id, s)
        
        return redirect(url_for("view_week", week_id=next_week.id))


@app.route("/export/excel/<int:week_id>")
def export_schedule_excel(week_id: int):
    """Export schedule to Excel file using the existing template"""
    import os
    
    with SessionLocal() as s:
        week = s.get(Week, week_id)
        if not week:
            return jsonify({"error": "Week not found"}), 404
        
        # Check if template exists
        template_path = "ScheduleTemplate.xlsx"
        if not os.path.exists(template_path):
            return jsonify({"error": "ScheduleTemplate.xlsx not found"}), 404
        
        # Get schedule data
        ctx = build_week_context(week_id)
        dates = ctx["week"]["dates"]
        sections = ctx["week"]["sections"]
        
        # Load the template
        wb = load_workbook(template_path)
        ws = wb.active

        # Update the sheet title with a safe, short title (Excel max 31 chars)
        def safe_sheet_title(title: str) -> str:
            invalid = set('[]:*?/\\')
            cleaned = ''.join(ch for ch in title if ch not in invalid)
            return cleaned[:31]

        ws.title = safe_sheet_title(format_week_label(week.start_date))
        
        # Update the dates in row 3 (columns E-K)
        for i, date_info in enumerate(dates):
            col_letter = get_column_letter(5 + i)  # Start from column E (5)
            cell_ref = f'{col_letter}3'
            ws[cell_ref] = week.start_date + timedelta(days=i)
            try:
                # Display as e.g., 21-Aug
                ws[cell_ref].number_format = 'dd-mmm'
            except Exception:
                pass
        
        # Update day headers in row 4 (columns E-K)
        day_headers = ["THU", "FRI", "SAT", "SUN", "MON", "TUE", "WED"]
        for i, day in enumerate(day_headers):
            col_letter = get_column_letter(5 + i)  # Start from column E (5)
            ws[f'{col_letter}4'] = day
        
        # Define employee row mappings based on template structure
        # Front Desk: rows 5-22 (names in column D)
        # Shuttle: rows 24-35 (names in column D) 
        # Breakfast Bar: rows 37-42 (names in column D)
        
        # Create mapping of employee names to their row numbers
        employee_rows = {}
        
        # Front Desk employees (rows 5-22)
        for row in range(5, 23):
            name_cell = ws[f'D{row}']
            if name_cell.value and isinstance(name_cell.value, str):
                name = name_cell.value.strip().upper()
                if name and name != "-" and name != "TBD":
                    # Remove asterisk if present
                    clean_name = name.replace("*", "").strip()
                    employee_rows[clean_name] = {"row": row, "section": "Front Desk"}
        
        # Shuttle employees (rows 24-35)
        for row in range(24, 36):
            name_cell = ws[f'D{row}']
            if name_cell.value and isinstance(name_cell.value, str):
                name = name_cell.value.strip().upper()
                if name and name != "-" and name != "TBD":
                    # Remove asterisk if present
                    clean_name = name.replace("*", "").strip()
                    employee_rows[clean_name] = {"row": row, "section": "Shuttle"}
        
        # Breakfast Bar employees (rows 37-42)
        for row in range(37, 43):
            name_cell = ws[f'D{row}']
            if name_cell.value and isinstance(name_cell.value, str):
                name = name_cell.value.strip().upper()
                if name and name != "-" and name != "TBD":
                    # Remove asterisk if present
                    clean_name = name.replace("*", "").strip()
                    employee_rows[clean_name] = {"row": row, "section": "Breakfast Bar"}
        
        # Fill in the shift data
        import re
        def _normalize_shift_display(raw: str) -> str:
            if not raw:
                return ""
            if raw == "Set":
                return "-"
            # Prefer time-only inside parentheses if present
            if "(" in raw and ")" in raw:
                start = raw.find("(") + 1
                end = raw.find(")", start)
                if end > start:
                    raw = raw[start:end]
            # Add spaces around en dash and hyphen
            raw = re.sub(r"\s*–\s*", " – ", raw)
            raw = re.sub(r"\s*-\s*", " - ", raw)
            # Lowercase AM/PM tokens (including attached to times)
            raw = re.sub(r"AM|PM", lambda m: m.group(0).lower(), raw)
            return raw

        # Maps of employee -> set of date ISO strings
        vacation_days = ctx.get("vacation_days", {})
        dismissed_days = ctx.get("dismissed_days", {})
        for section_name in ["Breakfast Bar", "Front Desk", "Shuttle"]:
            if section_name not in sections:
                continue
                
            section = sections[section_name]
            assignments = section["assignments"]
            
            for employee_name, employee_assignments in assignments.items():
                # Find the employee in our mapping
                employee_key = employee_name.upper()
                if employee_key in employee_rows and employee_rows[employee_key]["section"] == section_name:
                    row_num = employee_rows[employee_key]["row"]
                    
                    # Fill in the shifts for each day (columns E-K)
                    for i, date_info in enumerate(dates):
                        col_letter = get_column_letter(5 + i)  # Start from column E (5)
                        date_key = date_info["key"]
                        shift_value = employee_assignments[date_key]

                        # For time-off, show as request type based on Vacation toggle
                        if shift_value == TIME_OFF_LABEL:
                            # Follow main schedule display: REQ VAC only when dismissed+vacation, else REQ OFF
                            is_dismissed = date_key in (dismissed_days.get(employee_name, set()) or set())
                            is_vac = date_key in (vacation_days.get(employee_name, set()) or set())
                            ws[f'{col_letter}{row_num}'] = "REQ VAC" if (is_dismissed and is_vac) else "REQ OFF"
                            continue

                        # Format shift display: time-only and normalized dash spacing
                        shift_display = _normalize_shift_display(shift_value)
                        # Front Desk export requirement: remove :00 and use hyphen
                        if section_name == "Front Desk":
                            # collapse ":00am/pm" -> "am/pm"
                            shift_display = re.sub(r":00(?=(am|pm))", "", shift_display)
                            # use hyphen instead of en dash
                            shift_display = shift_display.replace("–", "-")
                            # normalize spacing around hyphen
                            shift_display = re.sub(r"\s*-\s*", " - ", shift_display)

                        # Set the cell value
                        ws[f'{col_letter}{row_num}'] = shift_display
        
        # Save to BytesIO
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        # Generate filename like: "Sept 18 - Oct 15.xlsx"
        def month_abbrev(d: date) -> str:
            abbr = d.strftime('%b')
            # Prefer "Sept" over "Sep" for September
            return 'Sept' if d.month == 9 and abbr == 'Sep' else abbr
        end_date = week.start_date + timedelta(days=6)
        filename = f"{month_abbrev(week.start_date)} {week.start_date.day} - {month_abbrev(end_date)} {end_date.day}.xlsx"
        
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )


if __name__ == "__main__":
    init_db_once()
    app.run(host="0.0.0.0", port=5008, debug=True)
