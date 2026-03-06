from __future__ import annotations

import copy
import io
from dataclasses import dataclass, field
from datetime import date
from typing import List, Set, Optional, Tuple

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter


# =========================
# Konfiguration / Styling
# =========================

GREEN_FILL = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")


# =========================
# Datenmodelle
# =========================

@dataclass
class Employee:
    name: str
    is_fachkraft: bool
    availability: List[bool]  # Länge 31
    min_services: int
    max_services: int
    block_preferences: Set[int]
    wants_8_block: bool

    # Laufzeit-Status
    assigned_count: int = 0
    current_streak: int = 0
    last_day_assigned: Optional[int] = None

    # Block-Reservierungen
    locked_work_days: Set[int] = field(default_factory=set)  # 1..31
    locked_free_days: Set[int] = field(default_factory=set)  # 1..31

    def validate(self) -> List[str]:
        errors = []
        if not self.name.strip():
            errors.append("Name fehlt.")
        if len(self.availability) != 31:
            errors.append(f"{self.name or 'Mitarbeiter'}: Verfügbarkeit muss 31 Tage enthalten.")
        if self.min_services < 0 or self.max_services < 0:
            errors.append(f"{self.name or 'Mitarbeiter'}: Min/Max-Dienste dürfen nicht negativ sein.")
        if self.max_services < self.min_services:
            errors.append(f"{self.name or 'Mitarbeiter'}: Max-Dienste kleiner als Min-Dienste.")
        if not self.block_preferences and not self.wants_8_block:
            errors.append(f"{self.name or 'Mitarbeiter'}: Blockwunsch ist ein Muss-Feld.")
        return errors


@dataclass
class DayRequirement:
    target: int
    minimum: int
    needs_fachkraft: bool = True
    exact_target: bool = False  # Wochenende: exakt 3


# =========================
# Grundregeln
# =========================

def requirement_for_day(day: int, month: int, year: int) -> DayRequirement:
    wd = date(year, month, day).weekday()  # Montag=0 ... Sonntag=6
    if wd >= 4:  # Freitag, Samstag, Sonntag
        return DayRequirement(target=3, minimum=3, needs_fachkraft=True, exact_target=True)
    return DayRequirement(target=3, minimum=2, needs_fachkraft=True, exact_target=False)


def count_fachkraft(employees: List[Employee], assigned_ids: List[int]) -> int:
    return sum(1 for idx in assigned_ids if employees[idx].is_fachkraft)


def get_locked_workers_for_day(employees: List[Employee], day: int) -> List[int]:
    return [i for i, emp in enumerate(employees) if day in emp.locked_work_days]


def get_block_patterns(emp: Employee) -> List[Tuple[str, List[str]]]:
    """
    Rückgabe in Prioritätsreihenfolge.
    Größere Wunschblöcke zuerst.
    8er-Wunsch wird als 4 + frei + 4 modelliert.
    """
    patterns: List[Tuple[str, List[str]]] = []

    if emp.wants_8_block:
        patterns.append(("8er", ["work", "work", "work", "work", "off", "work", "work", "work", "work"]))

    for block_len in sorted(emp.block_preferences, reverse=True):
        patterns.append((f"{block_len}er", ["work"] * block_len))

    return patterns


def can_start_block(emp: Employee, start_day: int, pattern: List[str]) -> bool:
    """
    Prüft nur mit Blick auf den Mitarbeiter selbst:
    - passt in den Monat
    - verfügbar an allen Arbeitstagen
    - kollidiert nicht mit reservierten Frei-/Arbeitstagen
    - verletzt nicht die 4er-Regel
    - überschreitet max Dienste nicht
    """
    end_day = start_day + len(pattern) - 1
    if end_day > 31:
        return False

    work_days_needed = 0
    planned_workdays = len(emp.locked_work_days)

    streak = emp.current_streak
    previous_day = start_day - 1

    # Wenn direkt an bestehende Reservierung anschließt, Streak simulieren.
    # Falls der Vortag nicht der letzte echte Dienst war, starten wir neu.
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
    """
    Prüft, ob das Reservieren dieses Blocks an irgendeinem betroffenen Arbeitstag
    die Tageskapazität überschreiten würde.
    Obergrenze hier: 3 pro Tag.
    """
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
    """
    Bewertet mögliche Blockstarts.
    Hoher Score = besser.
    """
    score = 0

    # Größere echte Arbeitsblöcke bevorzugen
    score += pattern.count("work") * 10

    # 8er-Wunsch sehr hoch priorisieren, wenn möglich
    if block_name == "8er":
        score += 40

    # Mitarbeitende unter Min-Diensten bevorzugen
    if len(emp.locked_work_days) < emp.min_services:
        score += 20

    # Fairness: weniger bisher geplante Tage = mehr Bonus
    score += max(0, 30 - len(emp.locked_work_days) * 2)

    # Fachkräfte leicht bevorzugen
    if emp.is_fachkraft:
        score += 5

    # Wenn ein Block genau heute weitergeführt wird, leicht bevorzugen
    if (day - 1) in emp.locked_work_days:
        score += 8

    return score


def find_best_block_start(
    employees: List[Employee],
    day: int,
    month: int,
    year: int,
    need_fachkraft_now: bool,
) -> Optional[Tuple[int, str, List[str]]]:
    """
    Sucht den besten Mitarbeiter + Block, der heute starten kann.
    """
    candidates: List[Tuple[int, int, str, List[str]]] = []

    for i, emp in enumerate(employees):
        if day in emp.locked_work_days or day in emp.locked_free_days:
            continue

        patterns = get_block_patterns(emp)

        for block_name, pattern in patterns:
            if not can_start_block(emp, day, pattern):
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
) -> List[int]:
    """
    Falls an einem Tag trotz Blocklogik noch Plätze fehlen, wird mit 1er-Blocks aufgefüllt.
    Auch das ist ein echter Block, nur Länge 1.
    """
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
            if not can_start_block(emp, day, pattern):
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
            f"Fallback an Tag {day}: {employees[emp_idx].name} wurde mit 1er-Block ergänzt, "
            f"weil kein passender Wunschblock mehr möglich war."
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


def generate_schedule(employees_input: List[Employee], month: int, year: int) -> Tuple[List[List[int]], List[str], List[Employee]]:
    """
    Gibt zurück:
    - assignments_by_day
    - warnings
    - final_employees_state
    """
    employees = copy.deepcopy(employees_input)
    assignments_by_day: List[List[int]] = [[] for _ in range(31)]
    warnings: List[str] = []

    for day in range(1, 32):
        req = requirement_for_day(day, month, year)

        # Bereits reservierte Arbeitstage
        assigned = get_locked_workers_for_day(employees, day)

        # Falls schon zu viele reserviert wären, Warnung
        if len(assigned) > req.target:
            warnings.append(
                f"Überbesetzung Tag {day}: {len(assigned)} Personen reserviert, Ziel {req.target}. "
                f"Bitte Eingaben prüfen."
            )

        # Bis Zielbesetzung mit echten Wunschblöcken auffüllen
        while len(assigned) < req.target:
            need_fk_now = req.needs_fachkraft and count_fachkraft(employees, assigned) == 0
            best = find_best_block_start(employees, day, month, year, need_fk_now)

            if best is None:
                break

            emp_idx, block_name, pattern = best
            lock_block(employees[emp_idx], day, pattern)
            assigned = get_locked_workers_for_day(employees, day)

            warnings.append(
                f"Blockstart Tag {day}: {employees[emp_idx].name} startet {block_name}-Block."
            )

        # Falls Wunschblöcke nicht reichen: mit 1er-Blöcken ergänzen
        if len(assigned) < req.target:
            assigned = fill_day_with_fallback_workers(employees, day, month, year, assigned, warnings)

        # Final für diesen Tag
        assigned = get_locked_workers_for_day(employees, day)

        # Sicherheit: nicht mehr als 3 ausgeben
        if len(assigned) > req.target:
            assigned = assigned[:req.target]

        assignments_by_day[day - 1] = assigned

        # Tageswarnungen
        if len(assigned) < req.minimum:
            warnings.append(
                f"Unterbesetzung Tag {day}: {len(assigned)} eingeplant, Minimum {req.minimum}."
            )

        if req.needs_fachkraft and count_fachkraft(employees, assigned) == 0:
            warnings.append(f"Keine Fachkraft Tag {day} eingeplant.")

        update_states_for_day(employees, day, assigned)

    # Monatsabschluss
    for emp in employees:
        if emp.assigned_count < emp.min_services:
            warnings.append(
                f"Min-Dienste nicht erreicht: {emp.name} hat {emp.assigned_count}, Minimum {emp.min_services}."
            )

        if not emp.block_preferences and not emp.wants_8_block:
            warnings.append(f"Blockwunsch fehlt: {emp.name}")

    return assignments_by_day, warnings, employees


# =========================
# Excel Export
# =========================

def build_excel(
    original_employees: List[Employee],
    final_employees: List[Employee],
    assignments_by_day: List[List[int]],
    warnings: List[str],
    month: int,
    year: int
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dienstplan"

    ws["A1"] = "Name"
    ws["A1"].font = Font(bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")

    for d in range(1, 32):
        col = get_column_letter(1 + d)  # B..AF
        ws[f"{col}1"] = d
        ws[f"{col}1"].font = Font(bold=True)
        ws[f"{col}1"].alignment = Alignment(horizontal="center")

    sum_col = get_column_letter(33)  # AG
    ws[f"{sum_col}1"] = "Summe"
    ws[f"{sum_col}1"].font = Font(bold=True)
    ws[f"{sum_col}1"].alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 20
    for d in range(1, 32):
        ws.column_dimensions[get_column_letter(1 + d)].width = 5
    ws.column_dimensions[sum_col].width = 10

    assigned_sets = [set(day_ids) for day_ids in assignments_by_day]

    for row_idx, emp in enumerate(original_employees, start=2):
        ws[f"A{row_idx}"] = emp.name

        service_count = 0
        emp_idx = row_idx - 2

        for day_idx0 in range(31):
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
        ws[f"{sum_col}{row_idx}"].alignment = Alignment(horizontal="center")

    ws2 = wb.create_sheet("Warnungen")
    ws2["A1"] = "Warnungen"
    ws2["A1"].font = Font(bold=True)
    ws2.column_dimensions["A"].width = 120

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
# UI Hilfsfunktionen
# =========================

def employee_from_form(prefix: str) -> Employee:
    availability = [st.session_state.get(f"{prefix}_day_{d}", False) for d in range(1, 32)]

    return Employee(
        name=st.session_state.get(f"{prefix}_name", "").strip(),
        is_fachkraft=st.session_state.get(f"{prefix}_fachkraft", False),
        availability=availability,
        min_services=int(st.session_state.get(f"{prefix}_min", 0)),
        max_services=int(st.session_state.get(f"{prefix}_max", 0)),
        block_preferences=set(st.session_state.get(f"{prefix}_blocks", [])),
        wants_8_block=st.session_state.get(f"{prefix}_wants8", False),
    )


def preset_employee(prefix: str, index: int) -> None:
    if st.session_state.get(f"{prefix}_name"):
        return

    example_names = [
        "Anna", "Ben", "Clara", "David", "Elif", "Farid",
        "Greta", "Hasan", "Iris", "Jonas", "Klara", "Luca"
    ]
    default_name = example_names[index] if index < len(example_names) else f"Mitarbeiter {index+1}"

    st.session_state[f"{prefix}_name"] = default_name
    st.session_state[f"{prefix}_fachkraft"] = index in (0, 2, 4, 6)
    st.session_state[f"{prefix}_min"] = 8
    st.session_state[f"{prefix}_max"] = 15
    st.session_state[f"{prefix}_blocks"] = [2]
    st.session_state[f"{prefix}_wants8"] = False
    for d in range(1, 32):
        st.session_state[f"{prefix}_day_{d}"] = True


# =========================
# Streamlit App
# =========================

st.set_page_config(page_title="Dienstplaner Nachtwache", layout="wide")
st.title("Dienstplaner Nachtwache")
st.caption("Blockbasierte Planung mit Pflicht-Berücksichtigung von Blockwünschen und 8er-Muster (4 + frei + 4).")

if "employee_count" not in st.session_state:
    st.session_state.employee_count = 4

with st.sidebar:
    st.header("Planungsrahmen")
    month = st.number_input("Monat", min_value=1, max_value=12, value=3, step=1)
    year = st.number_input("Jahr", min_value=2025, max_value=2100, value=2026, step=1)

    st.markdown("---")
    if st.button("Mitarbeiter hinzufügen"):
        st.session_state.employee_count += 1

    if st.button("Mitarbeiter entfernen") and st.session_state.employee_count > 1:
        st.session_state.employee_count -= 1

    if st.button("Beispieldaten laden"):
        count = max(st.session_state.employee_count, 6)
        st.session_state.employee_count = count
        for i in range(count):
            preset_employee(f"emp_{i}", i)
        st.success("Beispieldaten geladen.")

st.subheader("Mitarbeiterdaten")

for i in range(st.session_state.employee_count):
    prefix = f"emp_{i}"
    with st.expander(f"Mitarbeiter {i + 1}", expanded=(i < 2)):
        col1, col2 = st.columns([2, 1])

        with col1:
            st.text_input("Name", key=f"{prefix}_name")
        with col2:
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
            help="Mindestens eine Blockgröße auswählen oder 8er-Wunsch aktivieren."
        )

        st.checkbox(
            "8er-Block-Wunsch (4 Dienst + 1 frei + 4 Dienst)",
            key=f"{prefix}_wants8",
        )

        st.write("Verfügbarkeit (Tag 1 bis 31)")
        cols = st.columns(7)
        for d in range(1, 32):
            with cols[(d - 1) % 7]:
                st.checkbox(f"Tag {d}", value=True, key=f"{prefix}_day_{d}")

st.markdown("---")

if st.button("Dienstplan erstellen", type="primary"):
    employees = [employee_from_form(f"emp_{i}") for i in range(st.session_state.employee_count)]

    errors: List[str] = []
    for emp in employees:
        errors.extend(emp.validate())

    if errors:
        for err in errors:
            st.error(err)
    else:
        assignments_by_day, warnings, final_employees = generate_schedule(employees, int(month), int(year))

        excel_bytes = build_excel(
            original_employees=employees,
            final_employees=final_employees,
            assignments_by_day=assignments_by_day,
            warnings=warnings,
            month=int(month),
            year=int(year),
        )

        st.success("Dienstplan wurde erstellt.")

        col_left, col_right = st.columns([2, 1])

        with col_left:
            st.subheader("Zusammenfassung")
            for emp in final_employees:
                st.write(
                    f"**{emp.name}** — Geplant: {emp.assigned_count}, "
                    f"Min: {emp.min_services}, Max: {emp.max_services}, "
                    f"Fachkraft: {'Ja' if emp.is_fachkraft else 'Nein'}"
                )

        with col_right:
            st.download_button(
                label="Excel-Datei herunterladen",
                data=excel_bytes,
                file_name=f"dienstplan_{int(month):02d}_{int(year)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.subheader("Warnungen")
        if warnings:
            for warning in warnings:
                st.warning(warning)
        else:
            st.info("Keine Warnungen.")
