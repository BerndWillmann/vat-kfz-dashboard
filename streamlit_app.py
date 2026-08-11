import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import date

# ============================================================
# VAT KFZ Werkstaetten-Cockpit V2.2
# SAP-Servicemeldungen + Kosten/Stunden aus zweitem Tabellenblatt
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
GARANTIE_PREFIXE = ["GM", "G"]
F_PREFIXE = ["F"]
KRITISCHE_REPARATUR_TAGE = 10
LANG_LAEUFER_TAGE = 30
WARTUNG_VORSCHAU_TAGE = 30

ABGESCHLOSSEN_STATUS = [
    "AMER MAUF MMAB",
    "MAUF MMAB",
    "MAUF MMAB MMDR",
    "MAUF MIAB MMDR",
]

IN_ARBEIT_STATUS = [
    "AMER MAUF MIAR",
    "MAUF MIAR",
    "MAUF MIAR MMDR",
    "MAUF MIAR OFMA",
]

OFFEN_STATUS = ["MOFN"]

LOESCHVERMERK_STATUS = [
    "LÖVM MAUF MMAB",
    "LÖVM MMAB",
    "LOEVM MAUF MMAB",
    "LOEVM MMAB",
]

# -----------------------------
# Allgemeine Hilfsfunktionen
# -----------------------------
def clean_colname(col):
    return str(col).strip().replace("\n", " ").replace("  ", " ")


def normalize_status(value):
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().upper().split())


def normalize_key(value):
    """SAP-Auftragsnummer robust als Text normalisieren, auch wenn Excel Zahlen liefert."""
    if pd.isna(value):
        return ""
    try:
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
    except Exception:
        pass
    text = str(value).strip()
    if text.endswith(".0") and text.replace(".0", "").isdigit():
        text = text[:-2]
    return text


def find_column(df, candidates, contains=False):
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


def get_existing_column(df, name):
    return name if name in df.columns else None


def safe_to_datetime(series):
    if series is None:
        return pd.Series(pd.NaT)
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def safe_to_number(series):
    if series is None:
        return pd.Series(dtype="float64")
    # SAP/Excel kann Dezimalkomma, Tausenderpunkte oder Leerzeichen enthalten.
    s = series.astype(str).str.replace(" ", "", regex=False)
    s = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0)


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
    if starts_with_any(art, F_PREFIXE):
        return "F Meldung"
    return "Sonstige"


def classify_status(value):
    status = normalize_status(value)
    if not status:
        return "Unbekannt"
    if status in [normalize_status(s) for s in LOESCHVERMERK_STATUS]:
        return "Loeschvermerk"
    if "LÖVM" in status or "LOEVM" in status:
        return "Loeschvermerk"
    if status in [normalize_status(s) for s in ABGESCHLOSSEN_STATUS]:
        return "Abgeschlossen"
    if status in [normalize_status(s) for s in IN_ARBEIT_STATUS]:
        return "In Arbeit"
    if status in [normalize_status(s) for s in OFFEN_STATUS]:
        return "Offen"
    if "MMAB" in status:
        return "Abgeschlossen"
    if "MIAR" in status:
        return "In Arbeit"
    if "MOFN" in status:
        return "Offen"
    return "Unbekannt"


def status_ampel(days):
    if pd.isna(days):
        return "⚪ unbekannt"
    if days <= 5:
        return "🟢 0-5 Tage"
    if days <= 10:
        return "🟡 6-10 Tage"
    if days <= 30:
        return "🟠 11-30 Tage"
    return "🔴 >30 Tage"


def euro(value):
    try:
        return f"{float(value):,.0f} €".replace(",", ".")
    except Exception:
        return "0 €"


def hours(value):
    try:
        return f"{float(value):,.1f} h".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,0 h"


def make_download(df, filename):
    csv = df.to_csv(index=False, sep=";").encode("utf-8-sig")
    st.download_button(
        label=f"{filename} herunterladen",
        data=csv,
        file_name=filename,
        mime="text/csv",
    )

# -----------------------------
# Excel laden
# -----------------------------
def load_workbook(uploaded_file):
    if uploaded_file is not None:
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    else:
        default_file = "Servicemeldungen_VAT_KFZ.xlsx"
        try:
            xls = pd.ExcelFile(default_file, engine="openpyxl")
        except Exception:
            return None, None, []

    sheet_names = xls.sheet_names
    meldungen_df = pd.read_excel(xls, sheet_name=sheet_names[0], engine="openpyxl")
    kosten_df = None
    if len(sheet_names) > 1:
        kosten_df = pd.read_excel(xls, sheet_name=sheet_names[1], engine="openpyxl")
    return meldungen_df, kosten_df, sheet_names

# -----------------------------
# Meldungsdaten vorbereiten
# -----------------------------
def prepare_meldungen(df):
    df = df.copy()
    df.columns = [clean_colname(c) for c in df.columns]

    col_art = get_existing_column(df, "Meldungsart") or find_column(df, ["Meldungsart", "Art", "MArt"], contains=True)
    col_techplatz = get_existing_column(df, "Techn. Platz") or find_column(df, ["Technischer Platz", "Techn. Platz", "Techn Platz", "Equipment", "Objekt"], contains=True)
    col_equipment = get_existing_column(df, "Equipment") or find_column(df, ["Equipment"], contains=True)
    col_meldung = get_existing_column(df, "Meldung") or find_column(df, ["Meldung", "Meldungsnummer", "Ursprungsmeldung"], contains=True)
    col_auftrag = get_existing_column(df, "Auftrag") or find_column(df, ["Auftrag", "Auftragsnummer", "IH-Auftrag"], contains=True)
    col_status = get_existing_column(df, "Systemstatus") or find_column(df, ["Systemstatus", "Status"], contains=True)
    col_anwenderstatus = get_existing_column(df, "Anwenderstat.") or find_column(df, ["Anwenderstat.", "Anwenderstatus"], contains=True)
    col_meldedatum = get_existing_column(df, "Angelegt am") or find_column(df, ["Angelegt am", "Meldungsdatum", "Datum"], contains=True)
    col_faellig = get_existing_column(df, "Gew.Ende") or get_existing_column(df, "progn. Ende") or find_column(df, ["Gew.Ende", "progn. Ende", "Fälligkeit", "Endtermin"], contains=True)
    col_progn_ende = get_existing_column(df, "progn. Ende") or find_column(df, ["progn. Ende"], contains=True)
    col_abschluss = get_existing_column(df, "Abschlußdatum") or get_existing_column(df, "Abschlussdatum") or find_column(df, ["Abschlußdatum", "Abschlussdatum"], contains=True)
    col_prio = get_existing_column(df, "Priorität") or find_column(df, ["Priorität", "Prio"], contains=True)
    col_kurztext = get_existing_column(df, "Beschreibung") or find_column(df, ["Beschreibung", "Kurztext", "Text"], contains=True)
    col_arbeitsplatz = get_existing_column(df, "Verantw.ArbPl.") or find_column(df, ["Verantw.ArbPl.", "Arbeitsplatz"], contains=True)
    col_standort = get_existing_column(df, "Standort") or find_column(df, ["Standort"], contains=True)

    columns = {
        "art": col_art,
        "techplatz": col_techplatz,
        "equipment": col_equipment,
        "meldung": col_meldung,
        "auftrag": col_auftrag,
        "status": col_status,
        "anwenderstatus": col_anwenderstatus,
        "meldedatum": col_meldedatum,
        "faellig": col_faellig,
        "progn_ende": col_progn_ende,
        "abschluss": col_abschluss,
        "prio": col_prio,
        "kurztext": col_kurztext,
        "arbeitsplatz": col_arbeitsplatz,
        "standort": col_standort,
    }

    missing = [name for name, col in {
        "Meldungsart": col_art,
        "Technischer Platz / Equipment": col_techplatz,
        "Meldungsdatum": col_meldedatum,
        "Systemstatus": col_status,
        "Auftrag": col_auftrag,
    }.items() if col is None]

    if col_art is None:
        return df, columns, missing

    df["_Meldungsart"] = df[col_art].astype(str).str.strip().str.upper()
    df["_Kategorie"] = df["_Meldungsart"].apply(classify_meldungsart)
    df["_Ausgeschlossen"] = df["_Kategorie"].eq("Ausgeschlossen")

    df["_Meldungsdatum"] = safe_to_datetime(df[col_meldedatum]) if col_meldedatum else pd.NaT
    df["_Faelligkeit"] = safe_to_datetime(df[col_faellig]) if col_faellig else pd.NaT
    if col_progn_ende:
        df["_Faelligkeit"] = df["_Faelligkeit"].fillna(safe_to_datetime(df[col_progn_ende]))
    df["_Abschlussdatum"] = safe_to_datetime(df[col_abschluss]) if col_abschluss else pd.NaT

    today = pd.Timestamp(date.today())
    df["_Systemstatus_norm"] = df[col_status].apply(normalize_status) if col_status else ""
    df["_Statusgruppe"] = df[col_status].apply(classify_status) if col_status else "Unbekannt"
    df["_Loeschvermerk"] = df["_Statusgruppe"].eq("Loeschvermerk")
    df["_Abgeschlossen"] = df["_Statusgruppe"].eq("Abgeschlossen")
    df["_Offen_Status"] = df["_Statusgruppe"].isin(["Offen", "In Arbeit", "Unbekannt"])

    df["_Tage_offen"] = (today - df["_Meldungsdatum"]).dt.days
    df.loc[df["_Tage_offen"] < 0, "_Tage_offen"] = pd.NA
    df["_Dauer_bis_Abschluss"] = (df["_Abschlussdatum"] - df["_Meldungsdatum"]).dt.days
    df.loc[df["_Dauer_bis_Abschluss"] < 0, "_Dauer_bis_Abschluss"] = pd.NA

    # Wartungslogik V2.1 bleibt erhalten.
    df["_Wartung_geplant"] = (
        df["_Kategorie"].eq("WE/WK Wartung")
        & df["_Offen_Status"]
        & df["_Faelligkeit"].notna()
        & (df["_Faelligkeit"] >= today)
    )
    df["_Wartung_naechste_30_Tage"] = df["_Wartung_geplant"] & (df["_Faelligkeit"] <= today + pd.Timedelta(days=WARTUNG_VORSCHAU_TAGE))
    df["_Wartung_im_Verzug"] = (
        df["_Kategorie"].eq("WE/WK Wartung")
        & df["_Offen_Status"]
        & df["_Faelligkeit"].notna()
        & (df["_Faelligkeit"] < today)
    )
    df["_Wartung_ohne_Termin_offen"] = df["_Kategorie"].eq("WE/WK Wartung") & df["_Offen_Status"] & df["_Faelligkeit"].isna()

    df["_Steuerungsrelevant_offen"] = (
        (df["_Kategorie"].ne("WE/WK Wartung") & df["_Offen_Status"])
        | df["_Wartung_im_Verzug"]
        | df["_Wartung_ohne_Termin_offen"]
    )
    df["_In_Werkstatt"] = (
        df["_Statusgruppe"].isin(["Offen", "In Arbeit"])
        & (df["_Kategorie"].ne("WE/WK Wartung") | df["_Wartung_im_Verzug"] | df["_Wartung_ohne_Termin_offen"])
    )
    df["_MF_groesser_10_Tage_offen"] = df["_Kategorie"].eq("MF Reparatur") & df["_Offen_Status"] & (df["_Tage_offen"] > KRITISCHE_REPARATUR_TAGE)
    df["_Groesser_30_Tage_offen"] = df["_Steuerungsrelevant_offen"] & (df["_Tage_offen"] > LANG_LAEUFER_TAGE)
    df["_Ampel_Standzeit"] = df["_Tage_offen"].apply(status_ampel)

    if col_auftrag:
        df["_Auftrag_key"] = df[col_auftrag].apply(normalize_key)
    else:
        df["_Auftrag_key"] = ""

    return df, columns, missing

# -----------------------------
# Kostenblatt vorbereiten
# -----------------------------
def prepare_kosten(kosten_df):
    if kosten_df is None:
        return None, {}, None

    df = kosten_df.copy()
    df.columns = [clean_colname(c) for c in df.columns]

    col_auftrag = get_existing_column(df, "Auftrag") or find_column(df, ["Auftrag", "Auftragsnummer", "IH-Auftrag"], contains=True)
    col_auftragsart = get_existing_column(df, "Auftragsart") or find_column(df, ["Auftragsart"], contains=True)
    col_eckstart = get_existing_column(df, "Eckstartermin") or find_column(df, ["Eckstartermin", "Eckstarttermin", "Starttermin"], contains=True)
    col_kurztext = get_existing_column(df, "Kurztext") or find_column(df, ["Kurztext", "Beschreibung"], contains=True)
    col_hours = get_existing_column(df, "Istarbeit Sum") or find_column(df, ["Istarbeit Sum", "Ist Arbeit Sum", "Istarbeit", "IstArbeit", "Arbeit Sum"], contains=True)
    col_unit = get_existing_column(df, "Einheit Arbeit") or find_column(df, ["Einheit Arbeit", "Einheit"], contains=True)
    col_cost_actual = get_existing_column(df, "GesKosten Is") or get_existing_column(df, "GesKosten Ist") or find_column(df, ["GesKosten Is", "GesKosten Ist", "Gesamtkosten Ist", "Istkosten", "Kosten Ist"], contains=True)
    col_cost_plan = get_existing_column(df, "GesKosten P") or get_existing_column(df, "GesKosten Plan") or find_column(df, ["GesKosten P", "GesKosten Plan", "Plankosten", "Plan Kosten"], contains=True)
    col_currency = get_existing_column(df, "Währung") or find_column(df, ["Währung", "Currency"], contains=True)
    col_standortwert = get_existing_column(df, "Standortwert") or find_column(df, ["Standortwert", "Verrechnung", "Erlös", "Erloes"], contains=True)
    col_techplatz = get_existing_column(df, "Techn. Platz") or find_column(df, ["Techn. Platz", "Technischer Platz", "Techn Platz"], contains=True)
    col_arbeitsplatz = get_existing_column(df, "Verantw.ArbPl.") or find_column(df, ["Verantw.ArbPl.", "Arbeitsplatz"], contains=True)
    col_bezeichnung = get_existing_column(df, "Bezeichnung") or find_column(df, ["Bezeichnung"], contains=True)

    cols = {
        "auftrag": col_auftrag,
        "auftragsart": col_auftragsart,
        "eckstart": col_eckstart,
        "kurztext": col_kurztext,
        "hours": col_hours,
        "unit": col_unit,
        "cost_actual": col_cost_actual,
        "cost_plan": col_cost_plan,
        "currency": col_currency,
        "standortwert": col_standortwert,
        "techplatz": col_techplatz,
        "arbeitsplatz": col_arbeitsplatz,
        "bezeichnung": col_bezeichnung,
    }

    if col_auftrag is None:
        return df, cols, "Im Kostenblatt wurde keine Auftrag-Spalte erkannt."

    df["_Auftrag_key"] = df[col_auftrag].apply(normalize_key)
    df["_Kosten_Stunden"] = safe_to_number(df[col_hours]) if col_hours else 0
    df["_Kosten_Ist"] = safe_to_number(df[col_cost_actual]) if col_cost_actual else 0
    df["_Kosten_Plan"] = safe_to_number(df[col_cost_plan]) if col_cost_plan else 0
    df["_Standortwert"] = safe_to_number(df[col_standortwert]) if col_standortwert else 0
    df["_Verrechenbar"] = df["_Standortwert"] > 0

    agg = df.groupby("_Auftrag_key", dropna=False).agg(
        Kosten_Stunden=("_Kosten_Stunden", "sum"),
        Kosten_Ist=("_Kosten_Ist", "sum"),
        Kosten_Plan=("_Kosten_Plan", "sum"),
        Standortwert=("_Standortwert", "sum"),
        Kostenzeilen=("_Auftrag_key", "size"),
    ).reset_index()
    agg["Verrechenbar"] = agg["Standortwert"] > 0

    return agg, cols, None


def merge_costs(meldungen_df, kosten_agg):
    df = meldungen_df.copy()
    if kosten_agg is None:
        for col in ["Kosten_Stunden", "Kosten_Ist", "Kosten_Plan", "Standortwert", "Kostenzeilen"]:
            df[col] = 0
        df["Verrechenbar"] = False
        return df

    df = df.merge(kosten_agg, on="_Auftrag_key", how="left")
    for col in ["Kosten_Stunden", "Kosten_Ist", "Kosten_Plan", "Standortwert", "Kostenzeilen"]:
        df[col] = df[col].fillna(0)
    df["Verrechenbar"] = df["Verrechenbar"].fillna(False)
    return df

# -----------------------------
# Tabellenanzeige
# -----------------------------
def detail_columns(columns):
    cols = []
    for key in ["meldung", "auftrag", "art", "techplatz", "equipment", "kurztext", "status", "anwenderstatus", "prio", "meldedatum", "faellig", "progn_ende", "abschluss", "arbeitsplatz", "standort"]:
        col = columns.get(key)
        if col and col not in cols:
            cols.append(col)
    for col in ["_Kategorie", "_Statusgruppe", "_Tage_offen", "_Dauer_bis_Abschluss", "Kosten_Stunden", "Kosten_Ist", "Kosten_Plan", "Standortwert", "Verrechenbar", "_Ampel_Standzeit"]:
        if col not in cols:
            cols.append(col)
    return [c for c in cols if c is not None]


def show_table(df, columns, max_rows=None):
    cols = [c for c in detail_columns(columns) if c in df.columns]
    if max_rows:
        st.dataframe(df[cols].head(max_rows), use_container_width=True, hide_index=True)
    else:
        st.dataframe(df[cols], use_container_width=True, hide_index=True)

# -----------------------------
# Oberfläche
# -----------------------------
st.title("🔧 VAT KFZ Werkstätten-Cockpit V2.2")
st.caption("SAP-Servicemeldungen plus Kosten/Stunden je Auftrag, technischer Platz und Equipment")

with st.sidebar:
    st.header("Datenquelle")
    uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx"])
    st.info("Die App liest Blatt 1 als Servicemeldungen und Blatt 2 als Kosten/Stunden. Verbindung über Auftrag.")

    st.header("Regeln")
    st.write("**MM** und **LÖVM** werden ausgeschlossen")
    st.write("**MF** = wichtige Reparaturen")
    st.write("**Wartungen zählen erst als offen, wenn Gew.Ende erreicht/überschritten ist**")
    st.write("Kosten/Stunden werden über **Auftrag** zugespielt")

raw_df, kosten_raw_df, sheet_names = load_workbook(uploaded_file)
if raw_df is None:
    st.warning("Bitte eine Excel-Datei hochladen oder 'Servicemeldungen_VAT_KFZ.xlsx' ins Repository legen.")
    st.stop()

prepared_df, columns, missing = prepare_meldungen(raw_df)
kosten_agg, kosten_cols, kosten_error = prepare_kosten(kosten_raw_df)
prepared_df = merge_costs(prepared_df, kosten_agg)

if columns.get("art") is None:
    st.error("Die Spalte 'Meldungsart' konnte nicht erkannt werden.")
    st.write(list(raw_df.columns))
    st.stop()

if missing:
    st.warning("Einige wichtige Spalten wurden nicht eindeutig erkannt: " + ", ".join(missing))
if kosten_raw_df is None:
    st.info("Kein zweites Tabellenblatt für Kosten/Stunden gefunden. Kostenbereiche bleiben leer.")
elif kosten_error:
    st.warning(kosten_error)

relevant_df = prepared_df[(~prepared_df["_Ausgeschlossen"]) & (~prepared_df["_Loeschvermerk"])].copy()

# -----------------------------
# Filter
# -----------------------------
with st.sidebar:
    st.header("Filter")
    kategorien = sorted(relevant_df["_Kategorie"].dropna().unique().tolist())
    selected_kategorien = st.multiselect("Meldungskategorie", kategorien, default=kategorien)

    statusgruppen = sorted(relevant_df["_Statusgruppe"].dropna().unique().tolist())
    selected_status = st.multiselect("Statusgruppe", statusgruppen, default=statusgruppen)

    if columns.get("techplatz"):
        tech_values = sorted(relevant_df[columns["techplatz"]].dropna().astype(str).unique().tolist())
        selected_tech = st.multiselect("Technischer Platz", tech_values, default=[])
    else:
        selected_tech = []

    if columns.get("prio"):
        prio_values = sorted(relevant_df[columns["prio"]].dropna().astype(str).unique().tolist())
        selected_prio = st.multiselect("Priorität", prio_values, default=[])
    else:
        selected_prio = []

filtered_df = relevant_df[relevant_df["_Kategorie"].isin(selected_kategorien) & relevant_df["_Statusgruppe"].isin(selected_status)].copy()
if selected_tech and columns.get("techplatz"):
    filtered_df = filtered_df[filtered_df[columns["techplatz"]].astype(str).isin(selected_tech)]
if selected_prio and columns.get("prio"):
    filtered_df = filtered_df[filtered_df[columns["prio"]].astype(str).isin(selected_prio)]

# -----------------------------
# KPIs
# -----------------------------
offene_mf = int(filtered_df[(filtered_df["_Kategorie"] == "MF Reparatur") & filtered_df["_Offen_Status"]].shape[0])
critical_mf = int(filtered_df[filtered_df["_MF_groesser_10_Tage_offen"]].shape[0])
overdue_maintenance = int(filtered_df[filtered_df["_Wartung_im_Verzug"]].shape[0])
planned_maintenance = int(filtered_df[filtered_df["_Wartung_geplant"]].shape[0])
steuerungsrelevant_offen = int(filtered_df[filtered_df["_Steuerungsrelevant_offen"]].shape[0])
langlaeufer_30 = int(filtered_df[filtered_df["_Groesser_30_Tage_offen"]].shape[0])

if columns.get("techplatz"):
    fahrzeuge_in_werkstatt = int(filtered_df[filtered_df["_In_Werkstatt"]][columns["techplatz"]].dropna().astype(str).nunique())
else:
    fahrzeuge_in_werkstatt = 0

gesamtkosten = float(filtered_df["Kosten_Ist"].sum())
gesamtstunden = float(filtered_df["Kosten_Stunden"].sum())
standortwert = float(filtered_df["Standortwert"].sum())
verrechenbare_auftraege = int(filtered_df[filtered_df["Verrechenbar"]]["_Auftrag_key"].nunique())

k1, k2, k3, k4 = st.columns(4)
k1.metric("Offene MF-Reparaturen", offene_mf)
k2.metric("MF >10 Tage offen", critical_mf)
k3.metric("Wartungen im Verzug", overdue_maintenance)
k4.metric("Aktuell in Werkstatt", fahrzeuge_in_werkstatt)

k5, k6, k7, k8 = st.columns(4)
k5.metric("Geplante Wartungen", planned_maintenance)
k6.metric("Steuerungsrelevant offen", steuerungsrelevant_offen)
k7.metric(">30 Tage steuerungsrelevant offen", langlaeufer_30)
k8.metric("Istkosten gesamt", euro(gesamtkosten))

k9, k10, k11, k12 = st.columns(4)
k9.metric("Iststunden gesamt", hours(gesamtstunden))
k10.metric("Standortwert / Verrechnung", euro(standortwert))
k11.metric("Verrechenbare Aufträge", verrechenbare_auftraege)
k12.metric("Tabellenblätter", len(sheet_names))

st.divider()

# -----------------------------
# Tabs
# -----------------------------
tab0, tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "Werkstattleiter",
    "Übersicht",
    "Kritische MF-Reparaturen",
    "Wartungen",
    "Equipment-Analyse",
    "Standzeiten",
    "Kosten & Stunden",
    "Kostentreiber",
    "Detaildaten",
])

with tab0:
    st.subheader("Werkstattleiter-Ansicht")
    st.write("Fokus auf steuerungsrelevante offene Arbeiten, Kosten und Stunden.")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Kritische MF-Reparaturen")
        critical = filtered_df[filtered_df["_MF_groesser_10_Tage_offen"]].copy().sort_values("_Tage_offen", ascending=False)
        if critical.empty:
            st.success("Keine MF-Reparaturen über 10 Tage offen.")
        else:
            show_table(critical, columns, max_rows=20)
    with c2:
        st.markdown("### Wartungen im Verzug")
        overdue = filtered_df[filtered_df["_Wartung_im_Verzug"]].copy()
        if not overdue.empty:
            overdue["_Tage_im_Verzug"] = (pd.Timestamp(date.today()) - overdue["_Faelligkeit"]).dt.days
            overdue = overdue.sort_values("_Tage_im_Verzug", ascending=False)
            cols = [c for c in detail_columns(columns) + ["_Tage_im_Verzug"] if c in overdue.columns]
            st.dataframe(overdue[cols].head(20), use_container_width=True, hide_index=True)
        else:
            st.success("Keine überfälligen Wartungen gefunden.")

with tab1:
    st.subheader("Übersicht")
    c1, c2 = st.columns(2)
    with c1:
        cat_count = filtered_df.groupby("_Kategorie").size().reset_index(name="Anzahl")
        fig = px.bar(cat_count, x="_Kategorie", y="Anzahl", title="Anzahl Meldungen nach Kategorie", text="Anzahl")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        status_count = filtered_df.groupby(["_Kategorie", "_Statusgruppe"]).size().reset_index(name="Anzahl")
        fig = px.bar(status_count, x="_Kategorie", y="Anzahl", color="_Statusgruppe", title="Status nach Kategorie", text="Anzahl")
        st.plotly_chart(fig, use_container_width=True)
    if filtered_df["_Meldungsdatum"].notna().any():
        trend = filtered_df.dropna(subset=["_Meldungsdatum"]).copy()
        trend["Monat"] = trend["_Meldungsdatum"].dt.to_period("M").dt.to_timestamp()
        trend_group = trend.groupby(["Monat", "_Kategorie"]).size().reset_index(name="Anzahl")
        fig = px.bar(trend_group, x="Monat", y="Anzahl", color="_Kategorie", title="Meldungen pro Monat")
        st.plotly_chart(fig, use_container_width=True)

with tab2:
    st.subheader("Kritische MF-Reparaturen, länger als 10 Tage offen")
    critical = filtered_df[filtered_df["_MF_groesser_10_Tage_offen"]].copy().sort_values("_Tage_offen", ascending=False)
    st.metric("Kritische MF-Reparaturen", int(critical.shape[0]))
    if critical.empty:
        st.success("Keine offenen MF-Reparaturen über 10 Tage gefunden.")
    else:
        show_table(critical, columns)
        make_download(critical, "kritische_mf_reparaturen.csv")

with tab3:
    st.subheader("Wartungen")
    st.write("Wartungen mit offenem Status und Gew.Ende in der Zukunft werden als geplant dargestellt und nicht als Verzug gezählt.")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Wartungen im Verzug")
        overdue = filtered_df[filtered_df["_Wartung_im_Verzug"]].copy()
        if not overdue.empty:
            overdue["_Tage_im_Verzug"] = (pd.Timestamp(date.today()) - overdue["_Faelligkeit"]).dt.days
            overdue = overdue.sort_values("_Tage_im_Verzug", ascending=False)
            st.dataframe(overdue[[c for c in detail_columns(columns) + ["_Tage_im_Verzug"] if c in overdue.columns]], use_container_width=True, hide_index=True)
            make_download(overdue, "wartungen_im_verzug.csv")
        else:
            st.success("Keine überfälligen Wartungen gefunden.")
    with c2:
        st.markdown("### Geplante Wartungen")
        planned = filtered_df[filtered_df["_Wartung_geplant"]].copy().sort_values("_Faelligkeit", ascending=True)
        if not planned.empty:
            show_table(planned, columns, max_rows=100)
            make_download(planned, "geplante_wartungen.csv")
        else:
            st.info("Keine geplanten Wartungen gefunden.")

with tab4:
    st.subheader("Equipment-Analyse / Technischer Platz")
    if not columns.get("techplatz"):
        st.warning("Techn. Platz wurde nicht erkannt.")
    else:
        tech_col = columns["techplatz"]
        c1, c2 = st.columns(2)
        with c1:
            top_all = filtered_df.groupby(tech_col).size().reset_index(name="Anzahl").sort_values("Anzahl", ascending=False).head(20)
            fig = px.bar(top_all, x="Anzahl", y=tech_col, orientation="h", title="Top 20 technische Plätze nach Meldungen")
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            mf_df = filtered_df[filtered_df["_Kategorie"] == "MF Reparatur"]
            top_mf = mf_df.groupby(tech_col).size().reset_index(name="Anzahl_MF").sort_values("Anzahl_MF", ascending=False).head(20)
            fig = px.bar(top_mf, x="Anzahl_MF", y=tech_col, orientation="h", title="Top 20 technische Plätze nach MF-Reparaturen")
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)

with tab5:
    st.subheader("Standzeiten und Langläufer")
    stand = filtered_df[filtered_df["_Steuerungsrelevant_offen"]].copy().sort_values("_Tage_offen", ascending=False)
    if stand.empty:
        st.success("Keine steuerungsrelevanten offenen Meldungen gefunden.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            ampel = stand.groupby("_Ampel_Standzeit").size().reset_index(name="Anzahl")
            fig = px.pie(ampel, names="_Ampel_Standzeit", values="Anzahl", title="Standzeit-Ampel")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            stand_cat = stand.groupby(["_Kategorie", "_Ampel_Standzeit"]).size().reset_index(name="Anzahl")
            fig = px.bar(stand_cat, x="_Kategorie", y="Anzahl", color="_Ampel_Standzeit", title="Standzeit nach Kategorie", text="Anzahl")
            st.plotly_chart(fig, use_container_width=True)
        show_table(stand, columns, max_rows=100)
        make_download(stand, "steuerungsrelevante_offene_langlaeufer.csv")

with tab6:
    st.subheader("Kosten & Stunden")
    st.write("Kosten und Stunden werden aus dem zweiten Tabellenblatt über die Auftragsnummer verbunden.")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Istkosten", euro(filtered_df["Kosten_Ist"].sum()))
    c2.metric("Iststunden", hours(filtered_df["Kosten_Stunden"].sum()))
    c3.metric("Plankosten", euro(filtered_df["Kosten_Plan"].sum()))
    c4.metric("Standortwert / Verrechnung", euro(filtered_df["Standortwert"].sum()))

    by_cat = filtered_df.groupby("_Kategorie").agg(
        Istkosten=("Kosten_Ist", "sum"),
        Iststunden=("Kosten_Stunden", "sum"),
        Plankosten=("Kosten_Plan", "sum"),
        Standortwert=("Standortwert", "sum"),
        Auftraege=("_Auftrag_key", "nunique"),
    ).reset_index().sort_values("Istkosten", ascending=False)
    st.dataframe(by_cat, use_container_width=True, hide_index=True)
    if not by_cat.empty:
        fig = px.bar(by_cat, x="_Kategorie", y="Istkosten", title="Istkosten nach Meldungskategorie", text="Istkosten")
        st.plotly_chart(fig, use_container_width=True)
    make_download(by_cat, "kosten_stunden_nach_kategorie.csv")

with tab7:
    st.subheader("Kostentreiber")
    if not columns.get("techplatz"):
        st.warning("Techn. Platz wurde nicht erkannt.")
    else:
        tech_col = columns["techplatz"]
        by_tech = filtered_df.groupby(tech_col).agg(
            Istkosten=("Kosten_Ist", "sum"),
            Iststunden=("Kosten_Stunden", "sum"),
            Plankosten=("Kosten_Plan", "sum"),
            Standortwert=("Standortwert", "sum"),
            Auftraege=("_Auftrag_key", "nunique"),
            Meldungen=("_Meldungsart", "size"),
        ).reset_index()
        by_tech = by_tech.sort_values("Istkosten", ascending=False)
        st.markdown("### Top 20 technische Plätze nach Istkosten")
        top_cost = by_tech.head(20)
        fig = px.bar(top_cost, x="Istkosten", y=tech_col, orientation="h", title="Top 20 technische Plätze nach Istkosten")
        fig.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("### Top 20 technische Plätze nach Iststunden")
        top_hours = by_tech.sort_values("Iststunden", ascending=False).head(20)
        fig = px.bar(top_hours, x="Iststunden", y=tech_col, orientation="h", title="Top 20 technische Plätze nach Iststunden")
        fig.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig, use_container_width=True)

        st.dataframe(by_tech, use_container_width=True, hide_index=True)
        make_download(by_tech, "kostentreiber_technischer_platz.csv")

    st.markdown("### Verrechnung / Standortwert")
    ver = filtered_df[filtered_df["Standortwert"] > 0].copy().sort_values("Standortwert", ascending=False)
    if ver.empty:
        st.info("Keine Datensätze mit Standortwert > 0 gefunden.")
    else:
        show_table(ver, columns, max_rows=200)
        make_download(ver, "verrechnung_standortwert.csv")

with tab8:
    st.subheader("Detaildaten")
    show_table(filtered_df, columns)
    make_download(filtered_df, "vat_kfz_servicemeldungen_gefiltert.csv")

# -----------------------------
# Diagnosebereich
# -----------------------------
with st.expander("Erkannte Spalten / Diagnose"):
    st.write("Tabellenblätter:")
    st.write(sheet_names)
    st.write("Spalten Meldungsblatt:")
    st.json(columns)
    st.write("Spalten Kostenblatt:")
    st.json(kosten_cols)
    st.write("Alle Spalten im Meldungsblatt:")
    st.write(list(raw_df.columns))
    if kosten_raw_df is not None:
        st.write("Alle Spalten im Kostenblatt:")
        st.write(list(kosten_raw_df.columns))
    if columns.get("status"):
        st.write("Statusgruppen nach SAP-Systemstatus:")
        status_diag = prepared_df.groupby([columns["status"], "_Statusgruppe"]).size().reset_index(name="Anzahl")
        st.dataframe(status_diag, use_container_width=True, hide_index=True)
    st.write("Kosten-Merge Diagnose:")
    st.write({
        "Meldungen_gesamt": int(prepared_df.shape[0]),
        "Meldungen_mit_Auftrag": int((prepared_df["_Auftrag_key"] != "").sum()),
        "Meldungen_mit_Kostenzeile": int((prepared_df["Kostenzeilen"] > 0).sum()),
        "Summe_Istkosten": float(prepared_df["Kosten_Ist"].sum()),
        "Summe_Iststunden": float(prepared_df["Kosten_Stunden"].sum()),
        "Summe_Standortwert": float(prepared_df["Standortwert"].sum()),
    })
