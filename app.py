import streamlit as st
import pandas as pd
import json
import os
from datetime import date, timedelta
import calendar
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Config ──────────────────────────────────────────────────────────────────
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
RESPONSES_FILE = os.path.join(DATA_DIR, "responses.json")
TEAM_FILE = os.path.join(DATA_DIR, "team.json")
ADMIN_PIN = "icf2026"

# Paare: werden immer zusammen eingeplant
PAIRS = [("Andreas", "Claudia"), ("Jan M.", "Maria M.")]

# Gottesdienstgrössen
SLOT_SIZES = {"09:30": 6, "11:30": 4}

# Einsatz-Grenzen pro Monat
MAX_PER_MONTH = 4
TARGET_PER_MONTH = 3  # Ideal

os.makedirs(DATA_DIR, exist_ok=True)

TL_DEFAULT = ["Anne", "Gitta", "Kerstin", "Lena L.", "Nikita", "Rebecca"]
TEAM_DEFAULT = [
    "Andreas", "Claudia", "David", "Gundula", "Helena", "Hivin",
    "Isabella", "Jan M.", "Jan R.", "Jessie", "Jürgen",
    "Lara", "Laura", "Lena K.", "Maria", "Maria M.",
    "Melanie", "Nathalie", "Nicole", "Ralf", "Sarah", "Sofia",
    "Sophie", "Susanne", "Ute K.", "Ute S.", "Veronika",
]

GERMAN_MONTHS = [
    "", "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember"
]

COLORS = {
    "tl_header": "2C5F8A",
    "tl_fill": "D6E8F5",
    "regular_header": "1A6B3A",
    "regular_fill": "D6F0E1",
    "sunday_fill": "F0F0F0",
    "month_header": "1A3C5E",
    "time_fill": "E8E8E8",
}


# ── Data helpers ─────────────────────────────────────────────────────────────

def load_team():
    if os.path.exists(TEAM_FILE):
        with open(TEAM_FILE) as f:
            return json.load(f)
    return {"tl": TL_DEFAULT, "regular": TEAM_DEFAULT}


def save_team(team):
    with open(TEAM_FILE, "w") as f:
        json.dump(team, f, ensure_ascii=False, indent=2)


def load_responses():
    if os.path.exists(RESPONSES_FILE):
        with open(RESPONSES_FILE) as f:
            return json.load(f)
    return {}


def save_responses(responses):
    with open(RESPONSES_FILE, "w") as f:
        json.dump(responses, f, ensure_ascii=False, indent=2)


def get_sundays(year: int, month: int) -> list[date]:
    cal = calendar.monthcalendar(year, month)
    sundays = []
    for week in cal:
        if week[6] != 0:
            sundays.append(date(year, month, week[6]))
    return sundays


def response_key(year: int, month: int) -> str:
    return f"{year}-{month:02d}"


# ── Pages ─────────────────────────────────────────────────────────────────────

def page_survey():
    st.title("🌸 Welcome-Dienst Verfügbarkeiten eintragen")
    st.markdown("Bitte trage ein, an welchen Sonntagen du eingesetzt werden kannst. **Kein Login nötig.**")

    team = load_team()
    all_members = sorted(team["tl"] + team["regular"])

    col1, col2 = st.columns(2)
    with col1:
        name = st.selectbox("Dein Name", ["– bitte wählen –"] + all_members)
    with col2:
        today = date.today()
        next_month = today.replace(day=1) + timedelta(days=32)
        years = [today.year, today.year + 1]
        year = st.selectbox("Jahr", years, index=0 if next_month.year == today.year else 1)
        month = st.selectbox("Monat", list(range(1, 13)),
                             format_func=lambda m: GERMAN_MONTHS[m],
                             index=next_month.month - 1)

    if name == "– bitte wählen –":
        st.info("Wähle deinen Namen aus der Liste.")
        return

    is_tl = name in team["tl"]
    sundays = get_sundays(year, month)

    st.markdown(f"### {GERMAN_MONTHS[month]} {year}")

    if is_tl:
        st.info("Du bist Tagesleitung und kannst pro Gottesdienst angeben, ob du als **TL** und/oder als **Welcomer** verfügbar bist.")
    else:
        st.markdown("Wähle für jeden Sonntag aus, wann du Zeit hast:")

    selections = {}
    for sunday in sundays:
        ds = sunday.strftime("%d.%m.")
        st.markdown(f"**{ds}**")

        if is_tl:
            for slot, label in [("09:30", "09:30 Uhr"), ("11:30", "11:30 Uhr")]:
                cols = st.columns([2, 2, 2, 1])
                cols[0].markdown(f"*{label}*")
                as_tl = cols[1].checkbox("Als Tagesleitung", key=f"{ds}_{slot}_tl")
                as_welcome = cols[2].checkbox("Als Welcomer", key=f"{ds}_{slot}_w")
                if ds not in selections:
                    selections[ds] = {}
                selections[ds][slot] = as_welcome
                selections[ds][f"{slot}_tl"] = as_tl
        else:
            cols = st.columns([2, 2, 2])
            c930 = cols[0].checkbox("09:30 Uhr", key=f"{ds}_930")
            c1130 = cols[1].checkbox("11:30 Uhr", key=f"{ds}_1130")
            selections[ds] = {"09:30": c930, "11:30": c1130}

        st.markdown("")

    st.markdown("---")
    note = st.text_area("Anmerkungen (optional)", placeholder="z. B. 'Am 15.06. nur wenn nötig'")

    if st.button("✅ Verfügbarkeit speichern", type="primary"):
        responses = load_responses()
        key = response_key(year, month)
        if key not in responses:
            responses[key] = {}
        responses[key][name] = {"availability": selections, "note": note}
        save_responses(responses)
        st.success(f"Danke, {name}! Deine Verfügbarkeit wurde gespeichert.")
        st.balloons()


def page_overview():
    st.title("📋 Übersicht – Rückmeldungen")
    st.caption("Dieser Bereich ist nur für die Koordination. Bitte PIN eingeben.")

    if "admin_ok" not in st.session_state:
        st.session_state.admin_ok = False

    if not st.session_state.admin_ok:
        pin = st.text_input("Admin-PIN", type="password")
        if st.button("Einloggen"):
            if pin == ADMIN_PIN:
                st.session_state.admin_ok = True
                st.rerun()
            else:
                st.error("Falscher PIN.")
        return

    team = load_team()
    all_members = sorted(team["tl"] + team["regular"])
    responses = load_responses()

    today = date.today()
    next_month = today.replace(day=1) + timedelta(days=32)
    col1, col2 = st.columns(2)
    with col1:
        year = st.selectbox("Jahr", [today.year, today.year + 1],
                            index=0 if next_month.year == today.year else 1, key="ov_year")
    with col2:
        month = st.selectbox("Monat", list(range(1, 13)),
                             format_func=lambda m: GERMAN_MONTHS[m],
                             index=next_month.month - 1, key="ov_month")

    key = response_key(year, month)
    month_responses = responses.get(key, {})
    sundays = get_sundays(year, month)

    responded = set(month_responses.keys())
    not_responded = set(all_members) - responded

    col_a, col_b = st.columns(2)
    with col_a:
        st.metric("✅ Geantwortet", len(responded))
    with col_b:
        st.metric("⏳ Ausstehend", len(not_responded))

    if not_responded:
        with st.expander("Noch keine Rückmeldung von:"):
            for n in sorted(not_responded):
                role = "TL" if n in team["tl"] else "Welcome"
                st.write(f"- {n} ({role})")

    if not month_responses:
        st.info("Noch keine Rückmeldungen für diesen Monat.")
        return

    st.markdown(f"### Verfügbarkeitsmatrix – {GERMAN_MONTHS[month]} {year}")
    st.caption("TL = Tagesleitung · W = Welcomer · TL+W = beides · — = nicht verfügbar")

    rows = []
    for n in sorted(month_responses.keys()):
        row = {"Name": n, "Rolle": "TL" if n in team["tl"] else "Welcome"}
        avail = month_responses[n].get("availability", {})
        for sunday in sundays:
            ds = sunday.strftime("%d.%m.")
            for slot in ["09:30", "11:30"]:
                as_tl = avail.get(ds, {}).get(f"{slot}_tl", False)
                as_w = avail.get(ds, {}).get(slot, False)
                if as_tl and as_w:
                    val = "TL+W"
                elif as_tl:
                    val = "TL"
                elif as_w:
                    val = "W"
                else:
                    val = "—"
                row[f"{ds} {slot}"] = val
        row["Anmerkung"] = month_responses[n].get("note", "")
        rows.append(row)

    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.markdown("---")
    if st.button("📊 Dienstplan erstellen & Excel exportieren", type="primary"):
        schedule = build_schedule(year, month, month_responses, team, sundays)
        path = export_excel(year, month, schedule, sundays)
        with open(path, "rb") as f:
            st.download_button(
                "⬇️ Dienstplan herunterladen",
                data=f,
                file_name=os.path.basename(path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        st.success("Dienstplan erstellt!")
        display_schedule(schedule, sundays)


def page_team():
    st.title("👥 Team verwalten")

    if "admin_ok" not in st.session_state or not st.session_state.admin_ok:
        st.warning("Bitte zuerst im Übersicht-Tab einloggen.")
        return

    team = load_team()

    st.markdown("### Tagesleitung (TL)")
    st.caption("TL-Personen können sich in der Umfrage auch als Welcomer eintragen.")
    tl_str = st.text_area("TL-Personen (eine pro Zeile)",
                           value="\n".join(team["tl"]), height=180)

    st.markdown("### Welcome-Team (einfacher Dienst)")
    reg_str = st.text_area("Team-Mitglieder (eine pro Zeile)",
                           value="\n".join(team["regular"]), height=300)

    if st.button("💾 Speichern", type="primary"):
        team["tl"] = [n.strip() for n in tl_str.splitlines() if n.strip()]
        team["regular"] = [n.strip() for n in reg_str.splitlines() if n.strip()]
        save_team(team)
        st.success("Team gespeichert!")


# ── Schedule builder ──────────────────────────────────────────────────────────

def build_schedule(year, month, responses, team, sundays):
    """
    Regeln:
    - Jede TL-Person mind. 1× pro Monat, bevorzugt beide Slots am selben Sonntag
    - 09:30: 1 TL + 6 Welcomer | 11:30: 1 TL + 4 Welcomer
    - Max 4 Einsätze/Person/Monat, Ziel 2–3
    - Paare (Andreas+Claudia, Jan M.+Maria M.) immer gemeinsam einteilen
    """
    schedule = {}
    counts = {n: 0 for n in team["tl"] + team["regular"]}

    def avail_for(name, ds, slot, as_tl=False):
        key = f"{slot}_tl" if as_tl else slot
        return responses.get(name, {}).get("availability", {}).get(ds, {}).get(key, False)

    sunday_ds = [s.strftime("%d.%m.") for s in sundays]

    # ── Schritt 1: TL-Zuteilung pro Sonntag ─────────────────────────────────
    # Eine TL übernimmt den ganzen Sonntag (beide Slots).
    # Bevorzuge TLs die an diesem Sonntag für BEIDE Slots als TL verfügbar sind.
    # Jede TL muss mind. 1× vorkommen → unbesetzte TLs haben Vorrang.

    tl_assignments = {ds: None for ds in sunday_ds}
    tl_count = {n: 0 for n in team["tl"]}

    def tl_full_avail(name, ds):
        return avail_for(name, ds, "09:30", as_tl=True) and avail_for(name, ds, "11:30", as_tl=True)

    def tl_any_avail(name, ds):
        return avail_for(name, ds, "09:30", as_tl=True) or avail_for(name, ds, "11:30", as_tl=True)

    # Sonntage nach Anzahl verfügbarer TLs sortieren (engste zuerst)
    sorted_ds = sorted(sunday_ds, key=lambda ds: sum(1 for n in team["tl"] if tl_any_avail(n, ds)))

    for ds in sorted_ds:
        unassigned = [n for n in team["tl"] if tl_count[n] == 0 and tl_any_avail(n, ds)]
        all_avail = [n for n in team["tl"] if tl_any_avail(n, ds)]

        # Bevorzuge unbesetzte TLs mit voller Verfügbarkeit
        candidates = sorted(
            unassigned or all_avail,
            key=lambda n: (0 if tl_full_avail(n, ds) else 1, tl_count[n])
        )
        if candidates:
            chosen = candidates[0]
            tl_assignments[ds] = chosen
            tl_count[chosen] += 1
            counts[chosen] += 2  # Ganzer Sonntag = 2 Einsätze

    # ── Schritt 2: Welcomer pro Slot ────────────────────────────────────────
    for sunday in sundays:
        ds = sunday.strftime("%d.%m.")
        schedule[ds] = {}
        tl_name = tl_assignments.get(ds) or "– fehlt –"

        for slot in ["09:30", "11:30"]:
            target = SLOT_SIZES[slot]

            # Verfügbare Welcomer (unter Max-Grenze)
            pool = [
                n for n in team["regular"] + team["tl"]
                if counts.get(n, 0) < MAX_PER_MONTH and avail_for(n, ds, slot)
            ]

            chosen = []
            used = set()

            # Paare: beide verfügbar → gemeinsam einteilen
            # Nur einer verfügbar → diese Person einzeln in den Pool lassen
            for p1, p2 in PAIRS:
                if p1 in pool and p2 in pool and len(chosen) + 2 <= target:
                    chosen += [p1, p2]
                    used |= {p1, p2}

            # Restliche Plätze: einzelne Personen (inkl. Paar-Hälften ohne Partner)
            individuals = sorted(
                [n for n in pool if n not in used],
                key=lambda n: (1 if counts.get(n, 0) >= TARGET_PER_MONTH else 0, counts.get(n, 0))
            )
            chosen += individuals[: target - len(chosen)]

            for n in chosen:
                counts[n] = counts.get(n, 0) + 1

            schedule[ds][slot] = {"tl": tl_name, "team": chosen}

    return schedule


def display_schedule(schedule, sundays):
    st.markdown("### Vorschau Dienstplan")
    for sunday in sundays:
        ds = sunday.strftime("%d.%m.")
        st.markdown(f"**{ds}**")
        for slot in ["09:30", "11:30"]:
            slot_data = schedule.get(ds, {}).get(slot, {})
            tl = slot_data.get("tl", "—")
            members = ", ".join(slot_data.get("team", []))
            col1, col2 = st.columns([1, 3])
            col1.markdown(f"*{slot}*")
            col2.markdown(f"TL: **{tl}** | Team: {members or '—'}")


# ── Excel export ──────────────────────────────────────────────────────────────

def export_excel(year, month, schedule, sundays):
    wb = Workbook()
    ws = wb.active
    ws.title = f"{GERMAN_MONTHS[month]} {year}"

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def cell(row, col, value="", bold=False, bg=None, fg="000000", center=False, size=10):
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(name="Arial", bold=bold, color=fg, size=size)
        if bg:
            c.fill = PatternFill("solid", start_color=bg)
        if center:
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
        return c

    # Month header spanning all columns
    total_cols = 1 + len(sundays) * 2
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    c = ws.cell(row=1, column=1, value=f"Welcome-Dienstplan – {GERMAN_MONTHS[month]} {year}")
    c.font = Font(name="Arial", bold=True, color="FFFFFF", size=14)
    c.fill = PatternFill("solid", start_color=COLORS["month_header"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Date headers + time slot sub-headers
    col = 2
    for sunday in sundays:
        ds = sunday.strftime("%d.%m.")
        ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + 1)
        cell(2, col, ds, bold=True, bg=COLORS["sunday_fill"], center=True)
        cell(3, col, "09:30 Uhr", bold=True, bg=COLORS["time_fill"], center=True)
        cell(3, col + 1, "11:30 Uhr", bold=True, bg=COLORS["time_fill"], center=True)
        col += 2

    cell(2, 1, "", bg=COLORS["sunday_fill"])
    cell(3, 1, "Rolle", bold=True, bg=COLORS["time_fill"], center=True)
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 22

    # TL row
    row = 4
    cell(row, 1, "Tagesleitung", bold=True, bg=COLORS["tl_fill"],
         fg=COLORS["tl_header"], center=True)
    ws.row_dimensions[row].height = 20
    col = 2
    for sunday in sundays:
        ds = sunday.strftime("%d.%m.")
        for slot in ["09:30", "11:30"]:
            name = schedule.get(ds, {}).get(slot, {}).get("tl", "—")
            cell(row, col, name, bg=COLORS["tl_fill"], center=True)
            col += 1

    # Team member rows
    for i in range(1, 7):
        row = 4 + i
        cell(row, 1, str(i), bold=True, bg=COLORS["regular_fill"],
             fg=COLORS["regular_header"], center=True)
        ws.row_dimensions[row].height = 18
        col = 2
        for sunday in sundays:
            ds = sunday.strftime("%d.%m.")
            for slot in ["09:30", "11:30"]:
                members = schedule.get(ds, {}).get(slot, {}).get("team", [])
                name = members[i - 1] if i - 1 < len(members) else ""
                cell(row, col, name, center=True)
                col += 1

    ws.column_dimensions["A"].width = 16
    for i in range(2, 2 + len(sundays) * 2):
        ws.column_dimensions[get_column_letter(i)].width = 14

    path = os.path.join(DATA_DIR, f"Dienstplan_{GERMAN_MONTHS[month]}_{year}.xlsx")
    wb.save(path)
    return path


# ── Main ──────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="ICF Welcome – Dienstplan",
    page_icon="🙌",
    layout="wide",
)

st.sidebar.title("🫶 ICF Welcome-Team")
st.sidebar.markdown("---")

page = st.sidebar.radio(
    "Navigation",
    ["🌸 Verfügbarkeiten eintragen", "📋 Übersicht & Dienstplan", "👥 Team verwalten"],
)

if page == "🌸 Verfügbarkeiten eintragen":
    page_survey()
elif page == "📋 Übersicht & Dienstplan":
    page_overview()
elif page == "👥 Team verwalten":
    page_team()

st.sidebar.markdown("---")
st.sidebar.caption("ICF Welcome-Team Tool · v1.1")
