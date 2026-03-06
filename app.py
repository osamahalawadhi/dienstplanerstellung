from __future__ import annotations

import io
from dataclasses import dataclass
from datetime import date
from typing import List, Set, Optional, Tuple

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter


GREEN_FILL = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")


@dataclass
class Employee:
    name: str
    is_fachkraft: bool
    availability: List[bool]  # 31 Tage
    min_services: int
    max_services: int
    block_preferences: Set[int]
    wants_8_block: bool
    assigned_count: int = 0
    current_streak: int = 0
    last_day_assigned: Optional[int] = None


@dataclass
class DayRequirement:
    target: int
    minimum: int
    needs_fachkraft: bool = True
    exact_target: bool = False


def requirement_for_day(day: int, month: int, year: int) -> DayRequirement:
    wd = date(year, month, day).weekday()  # Mon=0 ... Sun=6
    if wd >= 4:  # Fr, Sa, So
        return DayRequirement(target=3, minimum=3, needs_fachkraft=True, exact_target=True)
    return DayRequirement(target=3, minimum=2, needs_fachkraft=True, exact_target=False)


def can_work(emp: Employee, day_idx0: int) -> bool:
    if not emp.availability[day_idx0]:
        return False
    if emp.assigned_count >= emp.max_services:
        return False
    if emp.current_streak >= 4:
        return False
    return True


def count_fachkraft(employees: List[Employee], assigned_ids: List[int]) -> int:
    return sum(1 for i in assigned_ids if employees[i].is_fachkraft)


def will_fit_block_start(emp: Employee, day_idx0: int, L: int) -> bool:
    if L < 1 or L > 4:
        return False
    if day_idx0 + (L - 1) > 30:
        return False
    for k in range(L):
        if not emp.availability[day_idx0 + k]:
            return False
    if emp.current_streak + L > 4:
        return False
    if emp.assigned_count + L > emp.max_services:
        return False
    return True


def block_fit_score(emp: Employee, day_idx0: int, worked_yesterday: bool) -> int:
    prefs = set(emp.block_preferences)
    if emp.wants_8_block:
        prefs.add(4)

    if not prefs:
        return 0

    if worked_yesterday:
        next_len = emp.current_streak + 1
        if next_len in prefs:
            return 40 + 5 * next_len
        return 10
    else:
        best = 0
        for L in sorted(prefs):
            if will_fit_block_start(emp, day_idx0, L):
                best = max(best, 20 + 5 * L)
        return best


def fairness_score(emp: Employee) -> int:
    score = max(0, 30 - emp.assigned_count * 2)
    if emp.assigned_count < emp.min_services:
        score += 15
    return score


def pick_day_assignments(employees: List[Employee], day_idx0: int, month: int, year: int):
    warnings = []
    req = requirement_for_day(day_idx0 + 1, month, year)

    candidates = [i for i, emp in enumerate(employees) if can_work(emp, day_idx0)]
    assigned = []

    def worked_yesterday(emp: Employee) -> bool:
        return emp.last_day_assigned == day_idx0

    for _ in range(req.target):
        if not candidates:
            break

        need_fk_now = req.needs_fachkraft and count_fachkraft(employees, assigned) == 0
        scored = []

        for i in candidates:
            emp = employees[i]
            fk_bonus = 50 if (need_fk_now and emp.is_fachkraft) else (0 if not need_fk_now else -100)
            total = (
                fk_bonus
                + block_fit_score(emp, day_idx0, worked_yesterday(emp))
                + fairness_score(emp)
                + (3 if emp.is_fachkraft else 0)
            )
            scored.append((total, i))

        scored.sort(reverse=True, key=lambda x: x[0])
        _, best_id = scored[0]
        assigned.append(best_id)
        candidates.remove(best_id)

    if len(assigned) < req.minimum:
        warnings.append(f"Unterbesetzung Tag {day_idx0 + 1}: {len(assigned)} eingeplant, Minimum {req.minimum}.")
    if req.needs_fachkraft and count_fachkraft(employees, assigned) == 0:
        warnings.append(f"Keine Fachkraft Tag {day_idx0 + 1} eingeplant.")

    return assigned, warnings


def update_states(employees: List[Employee], day_idx0: int, assigned_ids: List[int]) -> None:
    assigned_set = set(assigned_ids)
    day = day_idx0 + 1
    for i, emp in enumerate(employees):
        if i in assigned_set:
            emp.assigned_count += 1
            emp.current_streak += 1
            emp.last_day_assigned = day
        else:
            emp.current_streak = 0


def generate_schedule(employees: List[Employee], month: int, year: int):
    assignments_by_day = [[] for _ in range(31)]
    warnings = []

    for day_idx0 in range(31):
        assigned, day_warnings = pick_day_assignments(employees, day_idx0, month, year)
        assignments_by_day[day_idx0] = assigned
        warnings.extend(day_warnings)
        update_states(employees, day_idx0, assigned)

    for emp in employees:
        if emp.assigned_count < emp.min_services:
            warnings.append(
                f"Min-Dienste nicht erreicht: {emp.name} hat {emp.assigned_count}, Minimum {emp.min_services}."
            )

    return assignments_by_day, warnings


def build_excel(employees: List[Employee], assignments_by_day: List[List[int]], warnings: List[str], month: int, year: int) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dienstplan"

    ws["A1"] = "Name"
    ws["A1"].font = Font(bold=True)

    for d in range(1, 32):
        col = get_column_letter(1 + d)
        ws[f"{col}1"] = d
        ws[f"{col}1"].font = Font(bold=True)
        ws[f"{col}1"].alignment = Alignment(horizontal="center")

    sum_col = get_column_letter(33)
    ws[f"{sum_col}1"] = "Summe"
    ws[f"{sum_col}1"].font = Font(bold=True)

    ws.column_dimensions["A"].width = 20
    for d in range(1, 32):
        ws.column_dimensions[get_column_letter(1 + d)].width = 5
    ws.column_dimensions[sum_col].width = 10

    assigned_sets = [set(day_ids) for day_ids in assignments_by_day]

    for r, emp in enumerate(employees, start=2):
        ws[f"A{r}"] = emp.name
        service_count = 0

        for day_idx0 in range(31):
            c = get_column_letter(2 + day_idx0)
            cell = ws[f"{c}{r}"]
            cell.alignment = Alignment(horizontal="center")
            if (r - 2) in assigned_sets[day_idx0]:
                cell.value = "X"
                cell.fill = GREEN_FILL
                service_count += 1
            else:
                cell.value = ""

        ws[f"{sum_col}{r}"] = service_count

    ws2 = wb.create_sheet("Warnungen")
    ws2["A1"] = "Warnungen"
    ws2["A1"].font = Font(bold=True)
    ws2.column_dimensions["A"].width = 120

    for i, w in enumerate(warnings, start=2):
        ws2[f"A{i}"] = w

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def employee_from_form(prefix: str) -> Employee:
    name = st.session_state.get(f"{prefix}_name", "")
    is_fachkraft = st.session_state.get(f"{prefix}_fachkraft", False)
    min_services = st.session_state.get(f"{prefix}_min", 0)
    max_services = st.session_state.get(f"{prefix}_max", 0)
    blocks = set(st.session_state.get(f"{prefix}_blocks", []))
    wants_8 = st.session_state.get(f"{prefix}_wants8", False)
    availability = [st.session_state.get(f"{prefix}_day_{d}", False) for d in range(1, 32)]

    return Employee(
        name=name,
        is_fachkraft=is_fachkraft,
        availability=availability,
        min_services=min_services,
        max_services=max_services,
        block_preferences=blocks,
        wants_8_block=wants_8,
    )


st.set_page_config(page_title="Dienstplaner Nachtwache", layout="wide")
st.title("Dienstplaner Nachtwache")

if "employee_count" not in st.session_state:
    st.session_state.employee_count = 3

month = st.number_input("Monat", min_value=1, max_value=12, value=3, step=1)
year = st.number_input("Jahr", min_value=2025, max_value=2100, value=2026, step=1)

col_a, col_b = st.columns(2)
with col_a:
    if st.button("Mitarbeiter hinzufügen"):
        st.session_state.employee_count += 1
with col_b:
    if st.button("Mitarbeiter entfernen") and st.session_state.employee_count > 1:
        st.session_state.employee_count -= 1

st.markdown("---")

for i in range(st.session_state.employee_count):
    prefix = f"emp_{i}"
    with st.expander(f"Mitarbeiter {i + 1}", expanded=(i == 0)):
        st.text_input("Name", key=f"{prefix}_name")
        st.checkbox("Fachkraft", key=f"{prefix}_fachkraft")
        c1, c2 = st.columns(2)
        with c1:
            st.number_input("Min-Dienste", min_value=0, max_value=31, value=8, key=f"{prefix}_min")
        with c2:
            st.number_input("Max-Dienste", min_value=0, max_value=31, value=15, key=f"{prefix}_max")

        st.multiselect(
            "Bevorzugte Blockgrößen",
            options=[1, 2, 3, 4],
            default=[2],
            key=f"{prefix}_blocks",
        )
        st.checkbox("8er-Block-Wunsch (4 + frei + 4)", key=f"{prefix}_wants8")

        st.write("Verfügbarkeit")
        cols = st.columns(7)
        for d in range(1, 32):
            with cols[(d - 1) % 7]:
                st.checkbox(f"Tag {d}", value=True, key=f"{prefix}_day_{d}")

st.markdown("---")

if st.button("Dienstplan erstellen"):
    employees = []
    errors = []

    for i in range(st.session_state.employee_count):
        emp = employee_from_form(f"emp_{i}")
        if not emp.name.strip():
            errors.append(f"Mitarbeiter {i + 1}: Name fehlt.")
        if emp.max_services < emp.min_services:
            errors.append(f"{emp.name or f'Mitarbeiter {i + 1}'}: Max-Dienste kleiner als Min-Dienste.")
        if not emp.block_preferences and not emp.wants_8_block:
            errors.append(f"{emp.name or f'Mitarbeiter {i + 1}'}: Blockwunsch ist ein Muss-Feld.")
        employees.append(emp)

    if errors:
        for e in errors:
            st.error(e)
    else:
        assignments_by_day, warnings = generate_schedule(employees, int(month), int(year))
        excel_bytes = build_excel(employees, assignments_by_day, warnings, int(month), int(year))

        st.success("Dienstplan wurde erstellt.")

        if warnings:
            st.subheader("Warnungen")
            for w in warnings:
                st.warning(w)

        st.download_button(
            label="Excel-Datei herunterladen",
            data=excel_bytes,
            file_name=f"dienstplan_{int(month):02d}_{int(year)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
