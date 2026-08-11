import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import date

# ============================================================
# VAT KFZ Werkstaetten-Cockpit
# Datenquelle: SAP-Export als Excel-Datei
# Ziel: Werkstattauslastung, kritische Reparaturen, Wartungen im Verzug
# ============================================================

st.set_page_config(
    page_title="VAT KFZ Werkstaetten-Cockpit",
    page_icon="🔧",
    layout="wide"
)

# -----------------------------
# Einstellungen
# -----------------------------
AUSSCHLUSS_MELDUNGSARTEN = ["MM"]
REPARATUR_PREFIXE = ["MF"]
WARTUNG_PREFIXE = ["WE", "WK", "W1", "W2", "W3", "WZ", "WP"]
GARANTIE_PREFIXE = ["G", "GM"]
F_PREFIXE = ["F"]
KRITISCHE_REPARATUR_TAGE = 10

ABSCHLUSS_STATUS_WORTE = [
    "abgeschlossen",
    "arbeit erledigt",
    "erledigt",
    "geschlossen",
    "fertig",
    "storno"
]

# -----------------------------
# Hilfsfunktionen
# -----------------------------
def clean_colname(col):
    return str(col).strip().replace("\n", " ").replace("  ", " ")


def find_column(df, candidates, contains=False):
    """Findet eine passende Spalte, auch wenn SAP-Exporte leicht andere Namen haben."""
    cols_original = list(df.columns)
    cols_lower = {clean_colname(c).lower(): c for c in cols_original}

    for cand in candidates:
        cand_lower = cand.lower()
        if cand_lower in cols_lower:
            return cols_lower[cand_lower]

    if contains:
        for c in cols_original:
            c_lower = clean_colname(c).lower()
            for cand in candidates:
                if cand.lower() in c_lower:
                    return c
    return None


def safe_to_datetime(series):
    if series is None:
        return pd.Series(pd.NaT)
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def starts_with_any(value, prefixes):
    if pd.isna(value):
        return False
    value = str(value).strip().upper()
    return any(value.startswith(p.upper()) for p in prefixes)


def classify_meldungsart(value):
    art = "" if pd.isna(value) else str(value).strip().upper()

    if starts_with_any(art, AUSSCHLUSS_MELDUNGSARTEN):
        return "Ausgeschlossen"
    if starts_with_any(art, REPARATUR_PREFIXE):
        return "MF Reparatur"
    if starts_with_any(art, WARTUNG_PREFIXE):
        return "WE/WK Wartung"
    if starts_with_any(art, GARANTIE_PREFIXE):
        return "G Garantie"

    # Wichtig: MF ist bereits vorher abgefangen. Reines F wird separat dargestellt.
    if starts_with_any(art, F_PREFIXE):
        return "F Meldung"

    return "Sonstige"


def is_completed(row, status_col, completion_col):
    status_text = ""
    if status_col and not pd.isna(row.get(status_col)):
        status_text = str(row.get(status_col)).lower()

    status_done = any(word in status_text for word in ABSCHLUSS_STATUS_WORTE)

    completion_done = False
    if completion_col:
        completion_done = pd.notna(row.get(completion_col))

    return bool(status_done or completion_done)


def format_days(value):
    if pd.isna(value):
        return ""
    try:
        return f"{int(value)} Tage"
    except Exception:
        return str(value)


def status_ampel(days):
    if pd.isna(days):
        return "⚪ unbekannt"
    if days <= 5:
        return "🟢 0-5 Tage"
    if days <= 10:
        return "🟡 6-10 Tage"
    return "🔴 >10 Tage"


def load_excel(uploaded_file):
    if uploaded_file is not None:
        return pd.read_excel(uploaded_file, engine="openpyxl")

    # Optional: Wenn die Excel-Datei im GitHub-Repository liegt, wird sie automatisch geladen.
    default_file = "Servicemeldungen_VAT_KFZ.xlsx"
    try:
        return pd.read_excel(default_file, engine="openpyxl")
    except Exception:
        return None


def prepare_data(df):
    df = df.copy()
    df.columns = [clean_colname(c) for c in df.columns]

    col_art = find_column(df, ["Meldungsart", "Meldungsart Text", "Art", "Meld.Art", "MArt"], contains=True)
    col_techplatz = find_column(df, ["Technischer Platz", "Techn. Platz", "Techn Platz", "Equipment", "Equipement", "Objekt", "TechnPlatz"], contains=True)
    col_meldung = find_column(df, ["Meldung", "Meldungsnummer", "Servicemeldung", "Ursprungsmeldung", "Meldungsnr"], contains=True)
    col_status = find_column(df, ["Status", "Systemstatus", "Anwenderstatus", "Folgestatus"], contains=True)
    col_meldedatum = find_column(df, ["Meldungsdatum", "Angelegt am", "Erfassungsdatum", "Startdatum", "Meldedatum", "Datum"], contains=True)
    col_faellig = find_column(df, ["Fällig am", "Faellig am", "Fälligkeit", "Faelligkeit", "Endtermin", "Gewünschtes Ende", "Geplantes Ende", "Sollende"], contains=True)
    col_abschluss = find_column(df, ["Abschlussdatum", "Erledigt am", "Fertigstellung", "Rueckmeldedatum", "Rückmeldedatum", "Istende"], contains=True)
    col_prio = find_column(df, ["Priorität", "Prioritaet", "Prio"], contains=True)
    col_kurztext = find_column(df, ["Kurztext", "Beschreibung", "Meldungstext", "Text", "Schaden"], contains=True)
    col_auftrag = find_column(df, ["Auftrag", "Auftragsnummer", "IH-Auftrag"], contains=True)

    required = {
        "Meldungsart": col_art,
        "Technischer Platz / Equipment": col_techplatz,
        "Meldungsdatum": col_meldedatum,
        "Status": col_status
    }

    missing_required = [name for name, col in required.items() if col is None]

    if col_art is None:
        return df, {}, missing_required

    df["_Meldungsart"] = df[col_art].astype(str).str.strip().str.upper()
    df["_Kategorie"] = df["_Meldungsart"].apply(classify_meldungsart)
    df["_Ausgeschlossen"] = df["_Kategorie"].eq("Ausgeschlossen")

    if col_meldedatum:
        df["_Meldungsdatum"] = safe_to_datetime(df[col_meldedatum])
    else:
        df["_Meldungsdatum"] = pd.NaT

    if col_faellig:
        df["_Faelligkeit"] = safe_to_datetime(df[col_faellig])
    else:
        df["_Faelligkeit"] = pd.NaT

    if col_abschluss:
        df["_Abschlussdatum"] = safe_to_datetime(df[col_abschluss])
    else:
        df["_Abschlussdatum"] = pd.NaT

    today = pd.Timestamp(date.today())
    df["_Abgeschlossen"] = df.apply(lambda r: is_completed(r, col_status, "_Abschlussdatum"), axis=1)
    df["_Offen"] = ~df["_Abgeschlossen"]
    df["_Tage_offen"] = (today - df["_Meldungsdatum"]).dt.days
    df.loc[df["_Tage_offen"] < 0, "_Tage_offen"] = pd.NA

    df["_Wartung_im_Verzug"] = (
        df["_Kategorie"].eq("WE/WK Wartung")
        & df["_Offen"]
        & df["_Faelligkeit"].notna()
        & (df["_Faelligkeit"] < today)
    )

    df["_MF_groesser_10_Tage_offen"] = (
        df["_Kategorie"].eq("MF Reparatur")
        & df["_Offen"]
        & df["_Tage_offen"].notna()
        & (df["_Tage_offen"] > KRITISCHE_REPARATUR_TAGE)
    )

    df["_Ampel_Standzeit"] = df["_Tage_offen"].apply(status_ampel)

    columns = {
        "art": col_art,
        "techplatz": col_techplatz,
        "meldung": col_meldung,
        "status": col_status,
        "meldedatum": col_meldedatum,
        "faellig": col_faellig,
        "abschluss": col_abschluss,
        "prio": col_prio,
        "kurztext": col_kurztext,
        "auftrag": col_auftrag
    }
    return df, columns, missing_required


def detail_columns(columns):
    cols = []
    for key in ["meldung", "art", "techplatz", "kurztext", "status", "prio", "meldedatum", "faellig", "abschluss", "auftrag"]:
        col = columns.get(key)
        if col and col not in cols:
            cols.append(col)
    for col in ["_Kategorie", "_Tage_offen", "_Ampel_Standzeit"]:
        if col not in cols:
            cols.append(col)
    return cols

# -----------------------------
# Oberfläche
# -----------------------------
st.title("🔧 VAT KFZ Werkstätten-Cockpit")
st.caption("Streamlit-Dashboard für SAP-Servicemeldungen, Werkstattauslastung, Wartungen im Verzug und kritische MF-Reparaturen")

with st.sidebar:
    st.header("Datenquelle")
    uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx"])
    st.info("Tipp: Du kannst die Datei auch als 'Servicemeldungen_VAT_KFZ.xlsx' direkt ins GitHub-Repository legen.")

    st.header("Regeln")
    st.write("**MM** wird ausgeschlossen")
    st.write("**MF** = wichtige Reparaturen")
    st.write("**WE/WK/W1/W2/W3/WZ/WP** = Wartungen")
    st.write("**G/GM** = Garantie")
    st.write("**F** = eigene Kategorie")
    st.write(f"MF kritisch ab **>{KRITISCHE_REPARATUR_TAGE} Tagen offen**")

raw_df = load_excel(uploaded_file)

if raw_df is None:
    st.warning("Bitte eine Excel-Datei hochladen oder 'Servicemeldungen_VAT_KFZ.xlsx' ins Repository legen.")
    st.stop()

prepared_df, columns, missing = prepare_data(raw_df)

if "art" not in columns or columns.get("art") is None:
    st.error("Die Spalte 'Meldungsart' konnte nicht erkannt werden. Bitte prüfe den SAP-Export.")
    st.write("Gefundene Spalten:")
    st.write(list(raw_df.columns))
    st.stop()

if missing:
    st.warning("Einige wichtige Spalten wurden nicht eindeutig erkannt: " + ", ".join(missing))
    st.write("Die App läuft trotzdem weiter, manche Auswertungen können aber eingeschränkt sein.")

# MM ausschließen
relevant_df = prepared_df[~prepared_df["_Ausgeschlossen"]].copy()

# Sidebar-Filter
with st.sidebar:
    st.header("Filter")
    kategorien = sorted(relevant_df["_Kategorie"].dropna().unique().tolist())
    selected_kategorien = st.multiselect("Meldungskategorie", kategorien, default=kategorien)

    if columns.get("techplatz"):
        tech_values = sorted(relevant_df[columns["techplatz"]].dropna().astype(str).unique().tolist())
        selected_tech = st.multiselect("Technischer Platz / Equipment", tech_values, default=[])
    else:
        selected_tech = []

filtered_df = relevant_df[relevant_df["_Kategorie"].isin(selected_kategorien)].copy()
if selected_tech and columns.get("techplatz"):
    filtered_df = filtered_df[filtered_df[columns["techplatz"]].astype(str).isin(selected_tech)]

# -----------------------------
# KPI-Bereich
# -----------------------------
offene_mf = int(filtered_df[(filtered_df["_Kategorie"] == "MF Reparatur") & filtered_df["_Offen"]].shape[0])
critical_mf = int(filtered_df[filtered_df["_MF_groesser_10_Tage_offen"]].shape[0])
overdue_maintenance = int(filtered_df[filtered_df["_Wartung_im_Verzug"]].shape[0])
offen_total = int(filtered_df[filtered_df["_Offen"]].shape[0])
gesamt_relevant = int(filtered_df.shape[0])

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Offene MF-Reparaturen", offene_mf)
k2.metric("MF >10 Tage offen", critical_mf)
k3.metric("Wartungen im Verzug", overdue_maintenance)
k4.metric("Offene Meldungen gesamt", offen_total)
k5.metric("Relevante Meldungen ohne MM", gesamt_relevant)

st.divider()

# -----------------------------
# Tabs
# -----------------------------
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "Übersicht",
    "Kritische MF-Reparaturen",
    "Wartungen im Verzug",
    "Equipment-Analyse",
    "Standzeiten",
    "Detaildaten"
])

with tab1:
    st.subheader("Übersicht nach Meldungsart")
    c1, c2 = st.columns(2)

    with c1:
        cat_count = filtered_df.groupby("_Kategorie").size().reset_index(name="Anzahl")
        fig = px.bar(cat_count, x="_Kategorie", y="Anzahl", title="Anzahl Meldungen nach Kategorie", text="Anzahl")
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        open_count = filtered_df.groupby(["_Kategorie", "_Offen"]).size().reset_index(name="Anzahl")
        open_count["Status"] = open_count["_Offen"].map({True: "offen", False: "abgeschlossen"})
        fig = px.bar(open_count, x="_Kategorie", y="Anzahl", color="Status", title="Offen vs. abgeschlossen", text="Anzahl")
        st.plotly_chart(fig, use_container_width=True)

    if filtered_df["_Meldungsdatum"].notna().any():
        st.subheader("Meldungseingang nach Monat")
        trend = filtered_df.dropna(subset=["_Meldungsdatum"]).copy()
        trend["Monat"] = trend["_Meldungsdatum"].dt.to_period("M").dt.to_timestamp()
        trend_group = trend.groupby(["Monat", "_Kategorie"]).size().reset_index(name="Anzahl")
        fig = px.bar(trend_group, x="Monat", y="Anzahl", color="_Kategorie", title="Meldungen pro Monat")
        st.plotly_chart(fig, use_container_width=True)

with tab2:
    st.subheader("Kritische MF-Reparaturen, länger als 10 Tage offen")
    critical = filtered_df[filtered_df["_MF_groesser_10_Tage_offen"]].copy()
    critical = critical.sort_values("_Tage_offen", ascending=False)

    st.metric("Kritische MF-Reparaturen", int(critical.shape[0]))
    if critical.empty:
        st.success("Keine offenen MF-Reparaturen über 10 Tage gefunden.")
    else:
        st.dataframe(critical[detail_columns(columns)], use_container_width=True, hide_index=True)

with tab3:
    st.subheader("Wartungen im Verzug")
    overdue = filtered_df[filtered_df["_Wartung_im_Verzug"]].copy()
    overdue["_Tage_im_Verzug"] = (pd.Timestamp(date.today()) - overdue["_Faelligkeit"]).dt.days
    overdue = overdue.sort_values("_Tage_im_Verzug", ascending=False)

    st.metric("Wartungen im Verzug", int(overdue.shape[0]))
    if overdue.empty:
        st.success("Keine überfälligen Wartungen gefunden.")
    else:
        cols = detail_columns(columns) + ["_Tage_im_Verzug"]
        cols = [c for c in cols if c in overdue.columns]
        st.dataframe(overdue[cols], use_container_width=True, hide_index=True)

with tab4:
    st.subheader("Equipment-Analyse / Technischer Platz")

    if not columns.get("techplatz"):
        st.warning("Die Spalte 'Technischer Platz' bzw. 'Equipment' wurde nicht erkannt.")
    else:
        tech_col = columns["techplatz"]

        c1, c2 = st.columns(2)
        with c1:
            top_all = filtered_df.groupby(tech_col).size().reset_index(name="Anzahl").sort_values("Anzahl", ascending=False).head(20)
            fig = px.bar(top_all, x="Anzahl", y=tech_col, orientation="h", title="Top 20 Equipment nach allen relevanten Meldungen")
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            mf_df = filtered_df[filtered_df["_Kategorie"] == "MF Reparatur"]
            top_mf = mf_df.groupby(tech_col).size().reset_index(name="Anzahl_MF").sort_values("Anzahl_MF", ascending=False).head(20)
            fig = px.bar(top_mf, x="Anzahl_MF", y=tech_col, orientation="h", title="Top 20 Equipment nach MF-Reparaturen")
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("Equipment mit offenen Meldungen")
        open_by_equipment = filtered_df[filtered_df["_Offen"]].groupby(tech_col).agg(
            Offene_Meldungen=("_Offen", "size"),
            MF_offen=("_Kategorie", lambda x: (x == "MF Reparatur").sum()),
            Max_Tage_offen=("_Tage_offen", "max"),
            Durchschnitt_Tage_offen=("_Tage_offen", "mean")
        ).reset_index().sort_values(["MF_offen", "Max_Tage_offen"], ascending=False)

        st.dataframe(open_by_equipment, use_container_width=True, hide_index=True)

with tab5:
    st.subheader("Standzeiten und Langläufer")
    stand = filtered_df[filtered_df["_Offen"]].copy()
    stand = stand.sort_values("_Tage_offen", ascending=False)

    if stand.empty:
        st.success("Keine offenen Meldungen gefunden.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            ampel = stand.groupby("_Ampel_Standzeit").size().reset_index(name="Anzahl")
            fig = px.pie(ampel, names="_Ampel_Standzeit", values="Anzahl", title="Standzeit-Ampel offener Meldungen")
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            stand_cat = stand.groupby(["_Kategorie", "_Ampel_Standzeit"]).size().reset_index(name="Anzahl")
            fig = px.bar(stand_cat, x="_Kategorie", y="Anzahl", color="_Ampel_Standzeit", title="Standzeit nach Kategorie", text="Anzahl")
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("Längste offene Meldungen")
        st.dataframe(stand[detail_columns(columns)].head(100), use_container_width=True, hide_index=True)

with tab6:
    st.subheader("Detaildaten")
    st.write("Hier sind alle gefilterten, relevanten Meldungen ohne MM.")
    st.dataframe(filtered_df[detail_columns(columns)], use_container_width=True, hide_index=True)

    csv = filtered_df.to_csv(index=False, sep=";").encode("utf-8-sig")
    st.download_button(
        label="Gefilterte Daten als CSV herunterladen",
        data=csv,
        file_name="vat_kfz_servicemeldungen_gefiltert.csv",
        mime="text/csv"
    )

# -----------------------------
# Diagnosebereich
# -----------------------------
with st.expander("Erkannte Spalten / Diagnose"):
    st.write("Diese Spalten hat die App automatisch erkannt:")
    st.json(columns)
    st.write("Alle Spalten im Excel:")
    st.write(list(raw_df.columns))
