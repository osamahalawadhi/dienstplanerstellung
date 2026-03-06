import calendar
from datetime import date, datetime

import streamlit as st
from supabase import create_client


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
    return result.data


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
        "updated_at": datetime.utcnow().isoformat(),
    }

    if existing.data:
        return (
            sb.table("employee_inputs")
            .update(payload)
            .eq("id", existing.data[0]["id"])
            .execute()
        )

    return sb.table("employee_inputs").insert(payload).execute()


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

st.subheader("Bereits eingegangene Mitarbeitereingaben")
rows = load_employee_inputs(sb, planning_round_id)

if rows:
    for row in rows:
        updated_at = row.get("updated_at", "")
        submitted_text = "Ja" if row.get("submitted") else "Nein"
        st.write(f"**{row['name']}** — eingetragen: {submitted_text} — zuletzt geändert: {updated_at}")
else:
    st.info("Noch keine Mitarbeitereingaben vorhanden.")

st.markdown("---")
st.subheader("Eigene Daten eintragen")

with st.form("employee_form"):
    name = st.text_input("Name")
    is_fachkraft = st.checkbox("Fachkraft")
    c1, c2 = st.columns(2)
    with c1:
        min_services = st.number_input("Min-Dienste", min_value=0, max_value=31, value=8)
    with c2:
        max_services = st.number_input("Max-Dienste", min_value=0, max_value=31, value=15)

    block_preferences = st.multiselect(
        "Bevorzugte Blockgrößen",
        options=[1, 2, 3, 4],
        default=[2],
    )

    wants_8_block = st.checkbox("8er-Block-Wunsch (4 + frei + 4)")

    st.write("Verfügbarkeit")
    availability = []
    cols = st.columns(7)
    for d in range(1, days_in_month + 1):
        with cols[(d - 1) % 7]:
            available = st.checkbox(get_day_label(d, int(month), int(year)), value=True, key=f"day_{d}")
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
                wants_8_block=wants_8_block,
                availability=availability,
            )
            st.success("Deine Daten wurden gespeichert.")
            st.rerun()
        except Exception as e:
            st.error("Speichern fehlgeschlagen.")
            st.exception(e)