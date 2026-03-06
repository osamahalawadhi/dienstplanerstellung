import calendar
import copy
import io
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from typing import List, Set, Optional, Tuple

import streamlit as st
from supabase import create_client
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter


GREEN_FILL = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")


@dataclass
class Employee:
    name: str
    is_fachkraft: bool
    availability: List[bool]
    min_services: int
    max_services: int
    block_preferences: Set[int]
    wants_8_block: bool

    assigned_count: int = 0
    current_streak: int = 0
    last_day_assigned: Optional[int] = None

    locked_work_days: Set[int] = field(default_factory=set)
    locked_free_days: Set[int] = field(default_factory=set)


@dataclass
class DayRequirement:
    target: int
    minimum: int
    needs_fachkraft: bool = True
    exact_target: bool = False


@st.cache_resource
def get_supabase():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)


def get_days_in_month(month: int, year: int) -> int:
    return calendar.monthrange(year, month)[1]


def get_weekday_short(d: date) -> str:
    weekday_map = {
        0: "Mo",
        1: "Di",
        2: "Mi",
        3: "Do",
        4: "Fr",
        5: "Sa",
        6: "So",
    }
    return weekday_map[d.weekday()]


def get_day_label(day: int, month: int, year: int) -> str:
    d = date(year, month, day)
    return f"{get_weekday_short(d)} {d.strftime('%d.%m.%Y')}"


def get_excel_day_label(day: int, month: int, year: int) -> str:
    d = date(year, month, day)
    return f"{get_weekday_short(d)}\n{d.strftime('%d.%m.%Y')}"


def get_or_create_planning_round(sb, month: int, year: int):
    title = f"Dienstplan {month:02d}/{year}"

    existing = (
        sb.table("planning_rounds")
        .select("*")
        .eq("month", month)
        .eq("year", year)
        .execute()
    )

    if existing.data:
        return existing.data[0]

    created = (
        sb.table("planning_rounds")
        .insert({
            "month": month,
            "year": year,
            "title": title,
        })
        .execute()
    )
    return created.data[0]


def load_employee_inputs(sb, planning_round_id: int):
    result = (
        sb.table("employee_inputs")
        .select("*")
        .eq("planning_round_id", planning_round_id)
        .order("name")
        .execute()
    )
    return result.data or []


def load_employees_for_round(sb, planning_round_id: int):
    result = (
        sb.table("employees")
        .select("*")
        .eq("planning_round_id", planning_round_id)
        .eq("active", True)
        .order("name")
        .order("id")
        .execute()
    )
    rows = result.data or []

    unique_by_name = {}
    for row in rows:
        name = (row.get("name") or "").strip()
        if not name:
            continue
        if name not in unique_by_name:
            unique_by_name[name] = row

    return list(unique_by_name.values())


def load_existing_input_for_name(sb, planning_round_id: int, name: str):
    result = (
        sb.table("employee_inputs")
        .select("*")
        .eq("planning_round_id", planning_round_id)
        .eq("name", name)
        .limit(1)
        .execute()
    )

    if result.data:
        return result.data[0]
    return None


def save_employee_input(
    sb,
    planning_round_id: int,
    name: str,
    is_fachkraft: bool,
    min_services: int,
    max_services: int,
    block_preferences: list[int],
    wants_8_block: bool,
    availability: list[bool],
):
    existing = (
        sb.table("employee_inputs")
        .select("id")
        .eq("planning_round_id", planning_round_id)
        .eq("name", name)
        .execute()
    )

    payload = {
        "planning_round_id": planning_round_id,
        "name": name,
        "is_fachkraft": is_fachkraft,
        "min_services": min_services,
        "max_services": max_services,
        "block_preferences": block_preferences,
        "wants_8_block": wants_8_block,
        "availability": availability,
        "submitted": True,
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }

    if existing.data:
        return (
            sb.table("employee_inputs")
            .update(payload)
            .eq("id", existing.data[0]["id"])
            .execute()
        )

    return sb.table("employee_inputs").insert(payload).execute()


def build_employees_from_inputs(sb, planning_round_id: int, days_in_month: int) -> List[Employee]:
    rows = load_employee_inputs(sb, planning_round_id)

    employees = []
    for row in rows:
        if not row.get("submitted"):
            continue

        availability = row.get("availability") or []
        if len(availability) != days_in_month:
            availability = [False] * days_in_month

        blocks_raw = row.get("block_preferences") or []
        block_preferences = set()
        for b in blocks_raw:
            try:
                b_int = int(b)
                if b_int in {1, 2, 3, 4}:
                    block_preferences.add(b_int)
            except Exception:
                pass

        employees.append(
            Employee(
                name=row["name"],
                is_fachkraft=bool(row.get("is_fachkraft", False)),
                availability=[bool(x) for x in availability],
                min_services=int(row.get("min_services", 0)),
                max_services=int(row.get("max_services", 0)),
                block_preferences=block_preferences,
                wants_8_block=bool(row.get("wants_8_block", False)),
            )
        )

    return employees


def build_input_overview_excel(
    employees: List[Employee],
    month: int,
    year: int,
    days_in_month: int,
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Eingaben"

    headers = ["Name", "Fachkraft", "Min", "Max", "Blöcke", "8er-Wunsch"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for d in range(1, days_in_month + 1):
        col = get_column_letter(6 + d)
        ws[f"{col}1"] = get_day_label(d, month, year)
        ws[f"{col}1"].font = Font(bold=True)
        ws[f"{col}1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 12

    for d in range(1, days_in_month + 1):
        ws.column_dimensions[get_column_letter(6 + d)].width = 12

    ws.row_dimensions[1].height = 32

    for row_idx, emp in enumerate(employees, start=2):
        ws[f"A{row_idx}"] = emp.name
        ws[f"B{row_idx}"] = "Ja" if emp.is_fachkraft else "Nein"
        ws[f"C{row_idx}"] = emp.min_services
        ws[f"D{row_idx}"] = emp.max_services
        ws[f"E{row_idx}"] = ",".join(str(x) for x in sorted(emp.block_preferences))
        ws[f"F{row_idx}"] = "Ja" if emp.wants_8_block else "Nein"

        for day_idx, available in enumerate(emp.availability, start=1):
            col = get_column_letter(6 + day_idx)
            cell = ws[f"{col}{row_idx}"]
            cell.value = "Ja" if available else "Nein"
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if available:
                cell.fill = GREEN_FILL

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# =========================
# Scheduler-Logik
# =========================

def requirement_for_day(day: int, month: int, year: int) -> DayRequirement:
    wd = date(year, month, day).weekday()
    if wd >= 4:
        return DayRequirement(target=3, minimum=3, needs_fachkraft=True, exact_target=True)
    return DayRequirement(target=3, minimum=2, needs_fachkraft=True, exact_target=False)


def count_fachkraft(employees: List[Employee], assigned_ids: List[int]) -> int:
    return sum(1 for idx in assigned_ids if employees[idx].is_fachkraft)


def get_locked_workers_for_day(employees: List[Employee], day: int) -> List[int]:
    return [i for i, emp in enumerate(employees) if day in emp.locked_work_days]


def get_block_patterns(emp: Employee) -> List[Tuple[str, List[str]]]:
    patterns: List[Tuple[str, List[str]]] = []

    if emp.wants_8_block:
        patterns.append(("8er", ["work", "work", "work", "work", "off", "work", "work", "work", "work"]))

    for block_len in sorted(emp.block_preferences, reverse=True):
        patterns.append((f"{block_len}er", ["work"] * block_len))

    return patterns


def can_start_block(emp: Employee, start_day: int, pattern: List[str], days_in_month: int) -> bool:
    end_day = start_day + len(pattern) - 1
    if end_day > days_in_month:
        return False

    work_days_needed = 0
    planned_workdays = len(emp.locked_work_days)

    streak = emp.current_streak
    previous_day = start_day - 1

    if emp.last_day_assigned != previous_day and previous_day not in emp.locked_work_days:
        streak = 0

    for offset, token in enumerate(pattern):
        day = start_day + offset
        idx = day - 1

        if token == "work":
            if day in emp.locked_free_days:
                return False
            if day in emp.locked_work_days:
                return False
            if not emp.availability[idx]:
                return False

            work_days_needed += 1
            streak += 1
            if streak > 4:
                return False

        elif token == "off":
            if day in emp.locked_work_days:
                return False
            streak = 0

    if planned_workdays + work_days_needed > emp.max_services:
        return False

    return True


def block_respects_future_capacity(
    employees: List[Employee],
    start_day: int,
    pattern: List[str],
    month: int,
    year: int,
    emp_idx: int,
) -> bool:
    for offset, token in enumerate(pattern):
        if token != "work":
            continue

        day = start_day + offset
        req = requirement_for_day(day, month, year)
        already = get_locked_workers_for_day(employees, day)
        future_count = len(already) + (0 if emp_idx in already else 1)

        if future_count > req.target:
            return False

    return True


def lock_block(emp: Employee, start_day: int, pattern: List[str]) -> None:
    for offset, token in enumerate(pattern):
        day = start_day + offset
        if token == "work":
            emp.locked_work_days.add(day)
        elif token == "off":
            emp.locked_free_days.add(day)


def employee_priority_score(emp: Employee, block_name: str, pattern: List[str], day: int) -> int:
    score = 0
    score += pattern.count("work") * 10

    if block_name == "8er":
        score += 40

    if len(emp.locked_work_days) < emp.min_services:
        score += 20

    score += max(0, 30 - len(emp.locked_work_days) * 2)

    if emp.is_fachkraft:
        score += 5

    if (day - 1) in emp.locked_work_days:
        score += 8

    return score


def find_best_block_start(
    employees: List[Employee],
    day: int,
    month: int,
    year: int,
    need_fachkraft_now: bool,
    days_in_month: int,
) -> Optional[Tuple[int, str, List[str]]]:
    candidates: List[Tuple[int, int, str, List[str]]] = []

    for i, emp in enumerate(employees):
        if day in emp.locked_work_days or day in emp.locked_free_days:
            continue

        for block_name, pattern in get_block_patterns(emp):
            if not can_start_block(emp, day, pattern, days_in_month):
                continue
            if not block_respects_future_capacity(employees, day, pattern, month, year, i):
                continue

            score = employee_priority_score(emp, block_name, pattern, day)

            if need_fachkraft_now:
                if emp.is_fachkraft:
                    score += 100
                else:
                    score -= 100

            candidates.append((score, i, block_name, pattern))

    if not candidates:
        return None

    candidates.sort(key=lambda x: x[0], reverse=True)
    _, emp_idx, block_name, pattern = candidates[0]
    return emp_idx, block_name, pattern


def fill_day_with_fallback_workers(
    employees: List[Employee],
    day: int,
    month: int,
    year: int,
    assigned: List[int],
    warnings: List[str],
    days_in_month: int,
) -> List[int]:
    req = requirement_for_day(day, month, year)

    while len(assigned) < req.target:
        need_fk_now = req.needs_fachkraft and count_fachkraft(employees, assigned) == 0

        fallback_candidates: List[Tuple[int, int]] = []
        for i, emp in enumerate(employees):
            if i in assigned:
                continue
            if day in emp.locked_work_days or day in emp.locked_free_days:
                continue

            pattern = ["work"]
            if not can_start_block(emp, day, pattern, days_in_month):
                continue
            if not block_respects_future_capacity(employees, day, pattern, month, year, i):
                continue

            score = 0
            if len(emp.locked_work_days) < emp.min_services:
                score += 20
            score += max(0, 30 - len(emp.locked_work_days) * 2)
            if emp.is_fachkraft:
                score += 5

            if need_fk_now:
                if emp.is_fachkraft:
                    score += 100
                else:
                    score -= 100

            fallback_candidates.append((score, i))

        if not fallback_candidates:
            break

        fallback_candidates.sort(key=lambda x: x[0], reverse=True)
        _, emp_idx = fallback_candidates[0]
        lock_block(employees[emp_idx], day, ["work"])
        assigned = get_locked_workers_for_day(employees, day)
        warnings.append(
            f"Fallback an {get_day_label(day, month, year)}: {employees[emp_idx].name} "
            f"wurde mit 1er-Block ergänzt, weil kein passender Wunschblock mehr möglich war."
        )

    return assigned


def update_states_for_day(employees: List[Employee], day: int, assigned_ids: List[int]) -> None:
    assigned_set = set(assigned_ids)

    for i, emp in enumerate(employees):
        if i in assigned_set:
            emp.assigned_count += 1
            if emp.last_day_assigned == day - 1:
                emp.current_streak += 1
            else:
                emp.current_streak = 1
            emp.last_day_assigned = day
        else:
            emp.current_streak = 0


def generate_schedule(
    employees_input: List[Employee],
    month: int,
    year: int,
    days_in_month: int,
) -> Tuple[List[List[int]], List[str], List[Employee]]:
    employees = copy.deepcopy(employees_input)
    assignments_by_day: List[List[int]] = [[] for _ in range(days_in_month)]
    warnings: List[str] = []

    for day in range(1, days_in_month + 1):
        req = requirement_for_day(day, month, year)

        assigned = get_locked_workers_for_day(employees, day)

        if len(assigned) > req.target:
            warnings.append(
                f"Überbesetzung {get_day_label(day, month, year)}: "
                f"{len(assigned)} Personen reserviert, Ziel {req.target}."
            )

        while len(assigned) < req.target:
            need_fk_now = req.needs_fachkraft and count_fachkraft(employees, assigned) == 0
            best = find_best_block_start(employees, day, month, year, need_fk_now, days_in_month)

            if best is None:
                break

            emp_idx, block_name, pattern = best
            lock_block(employees[emp_idx], day, pattern)
            assigned = get_locked_workers_for_day(employees, day)

            warnings.append(
                f"Blockstart {get_day_label(day, month, year)}: "
                f"{employees[emp_idx].name} startet {block_name}-Block."
            )

        if len(assigned) < req.target:
            assigned = fill_day_with_fallback_workers(
                employees, day, month, year, assigned, warnings, days_in_month
            )

        assigned = get_locked_workers_for_day(employees, day)

        if len(assigned) > req.target:
            assigned = assigned[:req.target]

        assignments_by_day[day - 1] = assigned

        if len(assigned) < req.minimum:
            warnings.append(
                f"Unterbesetzung {get_day_label(day, month, year)}: "
                f"{len(assigned)} eingeplant, Minimum {req.minimum}."
            )

        if req.needs_fachkraft and count_fachkraft(employees, assigned) == 0:
            warnings.append(f"Keine Fachkraft an {get_day_label(day, month, year)} eingeplant.")

        update_states_for_day(employees, day, assigned)

    for emp in employees:
        if emp.assigned_count < emp.min_services:
            warnings.append(
                f"Min-Dienste nicht erreicht: {emp.name} hat {emp.assigned_count}, "
                f"Minimum {emp.min_services}."
            )

        if not emp.block_preferences and not emp.wants_8_block:
            warnings.append(f"Blockwunsch fehlt: {emp.name}")

    return assignments_by_day, warnings, employees


def build_schedule_excel(
    original_employees: List[Employee],
    final_employees: List[Employee],
    assignments_by_day: List[List[int]],
    warnings: List[str],
    month: int,
    year: int,
    days_in_month: int,
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dienstplan"

    ws["A1"] = "Name"
    ws["A1"].font = Font(bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    for d in range(1, days_in_month + 1):
        col = get_column_letter(1 + d)
        ws[f"{col}1"] = get_excel_day_label(d, month, year)
        ws[f"{col}1"].font = Font(bold=True)
        ws[f"{col}1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    sum_col = get_column_letter(days_in_month + 2)
    ws[f"{sum_col}1"] = "Summe"
    ws[f"{sum_col}1"].font = Font(bold=True)
    ws[f"{sum_col}1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions["A"].width = 20
    for d in range(1, days_in_month + 1):
        ws.column_dimensions[get_column_letter(1 + d)].width = 12
    ws.column_dimensions[sum_col].width = 10
    ws.row_dimensions[1].height = 32

    assigned_sets = [set(day_ids) for day_ids in assignments_by_day]

    for row_idx, emp in enumerate(original_employees, start=2):
        ws[f"A{row_idx}"] = emp.name
        emp_idx = row_idx - 2
        service_count = 0

        for day_idx0 in range(days_in_month):
            col = get_column_letter(2 + day_idx0)
            cell = ws[f"{col}{row_idx}"]
            cell.alignment = Alignment(horizontal="center", vertical="center")

            if emp_idx in assigned_sets[day_idx0]:
                cell.value = "X"
                cell.fill = GREEN_FILL
                service_count += 1
            else:
                cell.value = ""

        ws[f"{sum_col}{row_idx}"] = service_count
        ws[f"{sum_col}{row_idx}"].alignment = Alignment(horizontal="center", vertical="center")

    ws2 = wb.create_sheet("Warnungen")
    ws2["A1"] = "Warnungen"
    ws2["A1"].font = Font(bold=True)
    ws2.column_dimensions["A"].width = 140

    for i, warning in enumerate(warnings, start=2):
        ws2[f"A{i}"] = warning

    ws3 = wb.create_sheet("Statistik")
    headers = ["Name", "Fachkraft", "Min", "Max", "Geplant", "Wunschblöcke", "8er-Wunsch"]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws3.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True)

    for i, (orig, final) in enumerate(zip(original_employees, final_employees), start=2):
        ws3.cell(i, 1, orig.name)
        ws3.cell(i, 2, "Ja" if orig.is_fachkraft else "Nein")
        ws3.cell(i, 3, orig.min_services)
        ws3.cell(i, 4, orig.max_services)
        ws3.cell(i, 5, final.assigned_count)
        ws3.cell(i, 6, ",".join(map(str, sorted(orig.block_preferences))) if orig.block_preferences else "")
        ws3.cell(i, 7, "Ja" if orig.wants_8_block else "Nein")

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# =========================
# Streamlit UI
# =========================

st.set_page_config(page_title="Dienstplan Mitarbeitereingabe", layout="wide")
st.title("Dienstplan Mitarbeitereingabe")

sb = get_supabase()

with st.sidebar:
    st.header("Planungsmonat")
    month = st.number_input("Monat", min_value=1, max_value=12, value=3, step=1)
    year = st.number_input("Jahr", min_value=2025, max_value=2100, value=2026, step=1)
    days_in_month = get_days_in_month(int(month), int(year))

round_row = get_or_create_planning_round(sb, int(month), int(year))
planning_round_id = round_row["id"]

st.success(f"Planungsrunde aktiv: {round_row['title']}")

employees_master = load_employees_for_round(sb, planning_round_id)

if not employees_master:
    st.error("Keine Mitarbeitenden für diese Planungsrunde in der Tabelle 'employees' gefunden.")
    st.stop()

rows = load_employee_inputs(sb, planning_round_id)

submitted_by_name = {
    row["name"]: row
    for row in rows
    if row.get("submitted")
}

st.subheader("Status der Mitarbeitereingaben")

total_count = len(employees_master)
submitted_count = len(submitted_by_name)

st.info(f"{submitted_count} von {total_count} Mitarbeitenden haben bereits eingetragen.")

for emp in employees_master:
    name = emp["name"]
    if name in submitted_by_name:
        updated_at = submitted_by_name[name].get("updated_at", "")
        st.write(f"✅ **{name}** — eingetragen — zuletzt geändert: {updated_at}")
    else:
        st.write(f"❌ **{name}** — noch offen")

missing_names = [
    emp["name"]
    for emp in employees_master
    if emp["name"] not in submitted_by_name
]

if missing_names:
    st.warning("Noch offen: " + ", ".join(missing_names))
else:
    st.success("Alle Mitarbeitenden haben ihre Eingaben abgegeben.")

st.markdown("---")
st.subheader("Eigene Daten eintragen")

employee_options = [emp["name"] for emp in employees_master]
selected_name = st.selectbox("Mitarbeiter auswählen", employee_options)

selected_employee = next(emp for emp in employees_master if emp["name"] == selected_name)

name = selected_employee["name"]
is_fachkraft = selected_employee["is_fachkraft"]
min_services = selected_employee["min_services"]
max_services = selected_employee["max_services"]

st.info(
    f"Stammdaten für **{name}**: "
    f"Fachkraft: {'Ja' if is_fachkraft else 'Nein'}, "
    f"Min-Dienste: {min_services}, Max-Dienste: {max_services}"
)

existing_input = load_existing_input_for_name(sb, planning_round_id, selected_name)

default_blocks = [2]
default_wants_8 = False
default_availability = [True] * days_in_month

if existing_input:
    loaded_blocks = existing_input.get("block_preferences") or []
    loaded_wants_8 = existing_input.get("wants_8_block", False)
    loaded_availability = existing_input.get("availability") or []

    if isinstance(loaded_blocks, list):
        parsed_blocks = []
        for x in loaded_blocks:
            try:
                parsed_blocks.append(int(x))
            except Exception:
                pass
        default_blocks = [b for b in parsed_blocks if b in [1, 2, 3, 4]]

    default_wants_8 = bool(loaded_wants_8)

    if isinstance(loaded_availability, list) and len(loaded_availability) == days_in_month:
        default_availability = [bool(x) for x in loaded_availability]

with st.form("employee_form"):
    block_preferences = st.multiselect(
        "Bevorzugte Blockgrößen",
        options=[1, 2, 3, 4],
        default=default_blocks,
    )

    wants_8_block = st.checkbox(
        "8er-Block-Wunsch (4 + frei + 4)",
        value=default_wants_8,
    )

    st.write("Verfügbarkeit")
    availability = []
    cols = st.columns(7)

    for d in range(1, days_in_month + 1):
        with cols[(d - 1) % 7]:
            available = st.checkbox(
                get_day_label(d, int(month), int(year)),
                value=default_availability[d - 1],
                key=f"{selected_name}_day_{month}_{year}_{d}",
            )
            availability.append(available)

    submitted = st.form_submit_button("Speichern")

if submitted:
    errors = []

    if not name.strip():
        errors.append("Name fehlt.")
    if max_services < min_services:
        errors.append("Max-Dienste dürfen nicht kleiner als Min-Dienste sein.")
    if not block_preferences and not wants_8_block:
        errors.append("Mindestens ein Blockwunsch oder 8er-Wunsch ist erforderlich.")

    if errors:
        for err in errors:
            st.error(err)
    else:
        try:
            save_employee_input(
                sb=sb,
                planning_round_id=planning_round_id,
                name=name.strip(),
                is_fachkraft=is_fachkraft,
                min_services=int(min_services),
                max_services=int(max_services),
                block_preferences=list(block_preferences),
                wants_8_block=bool(wants_8_block),
                availability=availability,
            )
            st.success("Deine Daten wurden gespeichert.")
            st.rerun()
        except Exception as e:
            st.error("Speichern fehlgeschlagen.")
            st.exception(e)

st.markdown("---")
st.subheader("Admin-Bereich")

admin_mode = st.checkbox("Admin-Modus aktivieren")

if admin_mode:
    employees_for_plan = build_employees_from_inputs(sb, planning_round_id, days_in_month)

    st.write(f"Eingetragene Mitarbeitende für Planung: **{len(employees_for_plan)}**")

    if employees_for_plan:
        for emp in employees_for_plan:
            st.write(
                f"- {emp.name} | Fachkraft: {'Ja' if emp.is_fachkraft else 'Nein'} | "
                f"Min: {emp.min_services} | Max: {emp.max_services} | "
                f"Blöcke: {sorted(emp.block_preferences)} | 8er: {'Ja' if emp.wants_8_block else 'Nein'}"
            )

        overview_excel = build_input_overview_excel(
            employees_for_plan,
            int(month),
            int(year),
            days_in_month,
        )

        st.download_button(
            label="Eingaben als Kontroll-Excel herunterladen",
            data=overview_excel,
            file_name=f"eingaben_{int(month):02d}_{int(year)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown("---")
        if st.button("Dienstplan erstellen"):
            assignments_by_day, warnings, final_employees = generate_schedule(
                employees_for_plan,
                int(month),
                int(year),
                days_in_month,
            )

            schedule_excel = build_schedule_excel(
                original_employees=employees_for_plan,
                final_employees=final_employees,
                assignments_by_day=assignments_by_day,
                warnings=warnings,
                month=int(month),
                year=int(year),
                days_in_month=days_in_month,
            )

            st.success("Dienstplan wurde erstellt.")

            st.download_button(
                label="Dienstplan als Excel herunterladen",
                data=schedule_excel,
                file_name=f"dienstplan_{int(month):02d}_{int(year)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            st.subheader("Warnungen")
            if warnings:
                for warning in warnings:
                    st.warning(warning)
            else:
                st.info("Keine Warnungen.")
    else:
        st.warning("Noch keine vollständigen Mitarbeitereingaben vorhanden.")
