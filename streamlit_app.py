import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import date

# ============================================================
# VAT KFZ Werkstaetten-Cockpit V2.4
# SAP-Servicemeldungen + Kosten/Stunden aus zweitem Tabellenblatt
# VLAD = Leistung an Dritte = externe, verrechenbare Leistung / Erloese
# ============================================================

st.set_page_config(
    page_title="VAT KFZ Werkstaetten-Cockpit",
    page_icon="🔧",
    layout="wide"
)

# ------------------------------------------------------------
# Einstellungen
# ------------------------------------------------------------

AUSSCHLUSS_MELDUNGSARTEN = ["MM"]
REPARATUR_PREFIXE = ["MF"]
WARTUNG_PREFIXE = ["WE", "WK", "W1", "W2", "W3", "WZ", "WP"]
GARANTIE_PREFIXE = ["GM", "G"]
F_PREFIXE = ["F"]

KRITISCHE_REPARATUR_TAGE = 10
LANG_LAEUFER_TAGE = 30

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
    "LOEVM MAUF MMAB",
    "LOEVM MMAB",
    "LÖVM MAUF MMAB",
    "LÖVM MMAB",
]

# ------------------------------------------------------------
# Hilfsfunktionen
# ------------------------------------------------------------

def clean_colname(col):
    return str(col).strip().replace("\n", " ").replace("  ", " ")


def normalize_status(value):
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().upper().split())


def normalize_key(value):
    """SAP-Auftragsnummer robust als Text normalisieren."""
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
    if df is None:
        return None
    cols = list(df.columns)
    lower_map = {clean_colname(c).lower(): c for c in cols}
    for cand in candidates:
        key = cand.lower().strip()
        if key in lower_map:
            return lower_map[key]
    if contains:
        for cand in candidates:
            key = cand.lower().strip()
            for col in cols:
                if key in clean_colname(col).lower():
                    return col
    return None


def safe_to_datetime(series):
    if series is None:
        return pd.Series(pd.NaT)
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def safe_to_number(series):
    if series is None:
        return pd.Series(dtype="float64")
    s = series.astype(str).str.strip()
    s = s.str.replace(" ", "", regex=False)
    s = s.replace(["", "nan", "NaN", "None", "-", "–"], "0")

    # SAP/Excel in AT/DE: 1.234,56 -> 1234.56
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


def format_number_at(value, decimals=2):
    try:
        return f"{float(value):,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return f"{0:.{decimals}f}".replace(".", ",")


def euro(value):
    return format_number_at(value, 2) + " €"


def hours(value):
    return format_number_at(value, 1) + " h"


def percent(value):
    return format_number_at(value, 1) + " %"


def format_dataframe(df):
    if df is None:
        return df
    display_df = df.copy()
    euro_cols = [
        "Kosten_Ist", "Kosten_Plan", "Standortwert", "Interne_Kosten",
        "VLAD_Erloese", "Istkosten", "Plankosten", "VLAD_Erlöse",
        "Interne Kosten", "VLAD Erlöse"
    ]
    hour_cols = [
        "Kosten_Stunden", "Interne_Stunden", "VLAD_Stunden", "Iststunden",
        "Interne Stunden", "VLAD Stunden"
    ]
    for col in euro_cols:
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(euro)
    for col in hour_cols:
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(hours)
    return display_df


def display_dataframe(df, **kwargs):
    st.dataframe(format_dataframe(df), **kwargs)


def make_download(df, filename):
    csv = df.to_csv(index=False, sep=";").encode("utf-8-sig")
    st.download_button(
        label="CSV herunterladen",
        data=csv,
        file_name=filename,
        mime="text/csv",
    )

# ------------------------------------------------------------
# Excel laden
# ------------------------------------------------------------

@st.cache_data(show_spinner=False)
def load_workbook_from_file(uploaded_file):
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    sheet_names = xls.sheet_names
    raw_df = pd.read_excel(xls, sheet_name=sheet_names[0])
    kosten_df = None
    if len(sheet_names) > 1:
        kosten_df = pd.read_excel(xls, sheet_name=sheet_names[1])
    return raw_df, kosten_df, sheet_names


@st.cache_data(show_spinner=False)
def load_workbook_default():
    default_file = "Servicemeldungen_VAT_KFZ.xlsx"
    try:
        xls = pd.ExcelFile(default_file, engine="openpyxl")
        sheet_names = xls.sheet_names
        raw_df = pd.read_excel(xls, sheet_name=sheet_names[0])
        kosten_df = None
        if len(sheet_names) > 1:
            kosten_df = pd.read_excel(xls, sheet_name=sheet_names[1])
        return raw_df, kosten_df, sheet_names
    except Exception:
        return None, None, []


def load_workbook(uploaded_file):
    if uploaded_file is not None:
        return load_workbook_from_file(uploaded_file)
    return load_workbook_default()

# ------------------------------------------------------------
# Meldungsdaten vorbereiten
# ------------------------------------------------------------

def prepare_meldungen(df):
    df = df.copy()
    df.columns = [clean_colname(c) for c in df.columns]

    columns = {
        "meldung": find_column(df, ["Meldung", "Meldungsnummer", "Servicemeldung"], contains=True),
        "auftrag": find_column(df, ["Auftrag", "Auftragsnummer", "IH-Auftrag"], contains=True),
        "art": find_column(df, ["Meldungsart", "Art"], contains=True),
        "techplatz": find_column(df, ["Techn. Platz", "Technischer Platz", "Techplatz"], contains=True),
        "equipment": find_column(df, ["Equipment", "Equipmentnummer"], contains=True),
        "kurztext": find_column(df, ["Kurztext", "Beschreibung", "Meldungstext"], contains=True),
        "status": find_column(df, ["Systemstatus", "Status"], contains=True),
        "anwenderstatus": find_column(df, ["Anwenderstatus"], contains=True),
        "prio": find_column(df, ["Prioritaet", "Priorität", "Prio"], contains=True),
        "meldedatum": find_column(df, ["Meldedatum", "Meld.datum", "Erfassungsdatum"], contains=True),
        "faellig": find_column(df, ["Gewünschtes Ende", "Gewuenschtes Ende", "Gew. Ende", "Fälligkeit", "Faelligkeit"], contains=True),
        "progn_ende": find_column(df, ["Prognostiziertes Ende", "Progn. Ende", "Prognose Ende"], contains=True),
        "abschluss": find_column(df, ["Abschlussdatum", "Abgeschlossen am", "Abschluss"], contains=True),
        "arbeitsplatz": find_column(df, ["Arbeitsplatz", "Verantw. Arbeitsplatz", "Verantwortlicher Arbeitsplatz"], contains=True),
        "standort": find_column(df, ["Standort", "Werk"], contains=True),
    }

    missing = []
    for key in ["meldung", "auftrag", "art", "status"]:
        if columns.get(key) is None:
            missing.append(key)

    if columns["art"]:
        df["_Meldungsart"] = df[columns["art"]].astype(str).str.strip().str.upper()
    else:
        df["_Meldungsart"] = ""

    if columns["status"]:
        df["_Systemstatus_norm"] = df[columns["status"]].apply(normalize_status)
    else:
        df["_Systemstatus_norm"] = ""

    if columns["auftrag"]:
        df["_Auftrag_key"] = df[columns["auftrag"]].apply(normalize_key)
    else:
        df["_Auftrag_key"] = ""

    if columns["meldung"]:
        df["_Meldung_key"] = df[columns["meldung"]].apply(normalize_key)
    else:
        df["_Meldung_key"] = ""

    df["_Kategorie"] = df["_Meldungsart"].apply(classify_meldungsart)
    df["_Statusgruppe"] = df["_Systemstatus_norm"].apply(classify_status)
    df["_Ausgeschlossen"] = df["_Kategorie"] == "Ausgeschlossen"
    df["_Loeschvermerk"] = df["_Statusgruppe"] == "Loeschvermerk"
    df["_Offen_Status"] = df["_Statusgruppe"].isin(["Offen", "In Arbeit", "Unbekannt"])
    df["_Abgeschlossen_Status"] = df["_Statusgruppe"] == "Abgeschlossen"
    df["_In_Werkstatt"] = df["_Statusgruppe"].isin(["Offen", "In Arbeit"])

    df["_Meldungsdatum"] = safe_to_datetime(df[columns["meldedatum"]]) if columns["meldedatum"] else pd.NaT
    df["_Faelligkeit"] = safe_to_datetime(df[columns["faellig"]]) if columns["faellig"] else pd.NaT
    df["_Prognose_Ende"] = safe_to_datetime(df[columns["progn_ende"]]) if columns["progn_ende"] else pd.NaT
    df["_Abschlussdatum"] = safe_to_datetime(df[columns["abschluss"]]) if columns["abschluss"] else pd.NaT

    today = pd.Timestamp(date.today())
    df["_Tage_offen"] = (today - df["_Meldungsdatum"]).dt.days
    df.loc[df["_Abgeschlossen_Status"], "_Tage_offen"] = pd.NA
    df["_Dauer_bis_Abschluss"] = (df["_Abschlussdatum"] - df["_Meldungsdatum"]).dt.days
    df["_Ampel_Standzeit"] = df["_Tage_offen"].apply(status_ampel)

    df["_MF_groesser_10_Tage_offen"] = (
        (df["_Kategorie"] == "MF Reparatur")
        & df["_Offen_Status"]
        & (df["_Tage_offen"].fillna(0) > KRITISCHE_REPARATUR_TAGE)
    )

    df["_Groesser_30_Tage_offen"] = (
        df["_Offen_Status"]
        & (df["_Tage_offen"].fillna(0) > LANG_LAEUFER_TAGE)
    )

    df["_Wartung_im_Verzug"] = (
        (df["_Kategorie"] == "WE/WK Wartung")
        & df["_Offen_Status"]
        & df["_Faelligkeit"].notna()
        & (df["_Faelligkeit"] < today)
    )

    df["_Wartung_geplant"] = (
        (df["_Kategorie"] == "WE/WK Wartung")
        & df["_Offen_Status"]
        & df["_Faelligkeit"].notna()
        & (df["_Faelligkeit"] >= today)
    )

    df["_Steuerungsrelevant_offen"] = (
        df["_Offen_Status"]
        & (~df["_Ausgeschlossen"])
        & (~df["_Loeschvermerk"])
    )

    for col in [
        "Kosten_Stunden", "Kosten_Ist", "Kosten_Plan", "Standortwert", "Kostenzeilen",
        "Interne_Stunden", "Interne_Kosten", "VLAD_Stunden", "VLAD_Erloese"
    ]:
        df[col] = 0.0

    df["Auftragsart_Kostenblatt"] = ""
    df["Verrechenbar"] = False

    return df, columns, missing

# ------------------------------------------------------------
# Kostenblatt vorbereiten
# ------------------------------------------------------------

def prepare_kosten(kosten_df):
    if kosten_df is None:
        return None, {}, None

    df = kosten_df.copy()
    df.columns = [clean_colname(c) for c in df.columns]

    kosten_cols = {
        "auftrag": find_column(df, ["Auftrag", "Auftragsnummer", "IH-Auftrag"], contains=True),
        "auftragsart": find_column(df, ["Auftragsart", "AufArt"], contains=True),
        "kurztext": find_column(df, ["Kurztext", "Beschreibung"], contains=True),
        "istarbeit": find_column(df, ["Istarbeit", "Ist-Arbeit", "Arbeitszeit", "Stunden", "Iststunden"], contains=True),
        "kosten_ist": find_column(df, ["Kosten Ist", "Istkosten", "Ist-Kosten", "Kosten_Ist"], contains=True),
        "kosten_plan": find_column(df, ["Kosten Plan", "Plankosten", "Plan-Kosten", "Kosten_Plan"], contains=True),
        "standortwert": find_column(df, ["Standortwert", "Verrechnung", "Wert", "Erlös", "Erloes"], contains=True),
        "waehrung": find_column(df, ["Währung", "Waehrung", "Currency"], contains=True),
        "arbeitsplatz": find_column(df, ["Verantwortlicher Arbeitsplatz", "Arbeitsplatz", "Verantw. Arbeitsplatz"], contains=True),
    }

    if kosten_cols["auftrag"] is None:
        return None, kosten_cols, "Im zweiten Tabellenblatt wurde keine Auftragsnummer erkannt."

    df["_Auftrag_key"] = df[kosten_cols["auftrag"]].apply(normalize_key)

    if kosten_cols["auftragsart"]:
        df["Auftragsart_Kostenblatt"] = df[kosten_cols["auftragsart"]].astype(str).str.strip().str.upper()
    else:
        df["Auftragsart_Kostenblatt"] = ""

    df["Kosten_Stunden"] = safe_to_number(df[kosten_cols["istarbeit"]]) if kosten_cols["istarbeit"] else 0.0
    df["Kosten_Ist"] = safe_to_number(df[kosten_cols["kosten_ist"]]) if kosten_cols["kosten_ist"] else 0.0
    df["Kosten_Plan"] = safe_to_number(df[kosten_cols["kosten_plan"]]) if kosten_cols["kosten_plan"] else 0.0
    df["Standortwert"] = safe_to_number(df[kosten_cols["standortwert"]]) if kosten_cols["standortwert"] else 0.0

    df["Verrechenbar"] = df["Auftragsart_Kostenblatt"].eq("VLAD")
    df["Interne_Stunden"] = df.apply(lambda r: 0.0 if r["Verrechenbar"] else r["Kosten_Stunden"], axis=1)
    df["Interne_Kosten"] = df.apply(lambda r: 0.0 if r["Verrechenbar"] else r["Kosten_Ist"], axis=1)
    df["VLAD_Stunden"] = df.apply(lambda r: r["Kosten_Stunden"] if r["Verrechenbar"] else 0.0, axis=1)

    # Fachliche Dashboard-Sicht: VLAD = Leistung an Dritte.
    # Der Istkosten-Wert wird fuer VLAD als externer Erloes ausgewiesen.
    df["VLAD_Erloese"] = df.apply(lambda r: r["Kosten_Ist"] if r["Verrechenbar"] else 0.0, axis=1)

    agg = df.groupby("_Auftrag_key").agg(
        Kosten_Stunden=("Kosten_Stunden", "sum"),
        Kosten_Ist=("Kosten_Ist", "sum"),
        Kosten_Plan=("Kosten_Plan", "sum"),
        Standortwert=("Standortwert", "sum"),
        Interne_Stunden=("Interne_Stunden", "sum"),
        Interne_Kosten=("Interne_Kosten", "sum"),
        VLAD_Stunden=("VLAD_Stunden", "sum"),
        VLAD_Erloese=("VLAD_Erloese", "sum"),
        Verrechenbar=("Verrechenbar", "max"),
        Kostenzeilen=("_Auftrag_key", "size"),
    ).reset_index()

    art_agg = df.groupby("_Auftrag_key")["Auftragsart_Kostenblatt"].apply(
        lambda x: ", ".join(sorted(set([str(v) for v in x if str(v).strip() != ""])))
    ).reset_index()
    agg = agg.merge(art_agg, on="_Auftrag_key", how="left")

    return agg, kosten_cols, None


def merge_costs(meldungen_df, kosten_agg):
    df = meldungen_df.copy()
    if kosten_agg is None:
        return df

    # Vorinitialisierte Kostenfelder entfernen, damit nach Merge keine _x/_y-Spalten entstehen.
    drop_cols = [
        "Kosten_Stunden", "Kosten_Ist", "Kosten_Plan", "Standortwert", "Kostenzeilen",
        "Interne_Stunden", "Interne_Kosten", "VLAD_Stunden", "VLAD_Erloese",
        "Auftragsart_Kostenblatt", "Verrechenbar"
    ]
    df = df.drop(columns=[c for c in drop_cols if c in df.columns], errors="ignore")

    df = df.merge(kosten_agg, on="_Auftrag_key", how="left")

    numeric_cols = [
        "Kosten_Stunden", "Kosten_Ist", "Kosten_Plan", "Standortwert", "Kostenzeilen",
        "Interne_Stunden", "Interne_Kosten", "VLAD_Stunden", "VLAD_Erloese"
    ]
    for col in numeric_cols:
        if col not in df.columns:
            df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    if "Auftragsart_Kostenblatt" not in df.columns:
        df["Auftragsart_Kostenblatt"] = ""
    else:
        df["Auftragsart_Kostenblatt"] = df["Auftragsart_Kostenblatt"].fillna("")

    if "Verrechenbar" not in df.columns:
        df["Verrechenbar"] = False
    else:
        df["Verrechenbar"] = df["Verrechenbar"].fillna(False).astype(bool)

    return df

# ------------------------------------------------------------
# Tabellenanzeige
# ------------------------------------------------------------

def detail_columns(columns):
    cols = []
    for key in [
        "meldung", "auftrag", "art", "techplatz", "equipment", "kurztext", "status",
        "anwenderstatus", "prio", "meldedatum", "faellig", "progn_ende",
        "abschluss", "arbeitsplatz", "standort"
    ]:
        col = columns.get(key)
        if col and col not in cols:
            cols.append(col)

    add_cols = [
        "_Kategorie", "_Statusgruppe", "_Tage_offen", "_Dauer_bis_Abschluss",
        "Auftragsart_Kostenblatt", "Kosten_Stunden", "Kosten_Ist", "Kosten_Plan",
        "Interne_Stunden", "Interne_Kosten", "VLAD_Stunden", "VLAD_Erloese",
        "Standortwert", "Verrechenbar", "_Ampel_Standzeit"
    ]
    for col in add_cols:
        if col not in cols:
            cols.append(col)
    return [c for c in cols if c is not None]


def show_table(df, columns, max_rows=None):
    cols = [c for c in detail_columns(columns) if c in df.columns]
    if max_rows:
        display_dataframe(df[cols].head(max_rows), use_container_width=True, hide_index=True)
    else:
        display_dataframe(df[cols], use_container_width=True, hide_index=True)

# ------------------------------------------------------------
# Oberflaeche
# ------------------------------------------------------------

st.title("🔧 VAT KFZ Werkstaetten-Cockpit V2.4")
st.caption(
    "SAP-Servicemeldungen plus Kosten/Stunden je Auftrag. "
    "VLAD wird als Leistung an Dritte getrennt als externe Stunden und Erloese ausgewiesen."
)

with st.sidebar:
    st.header("Datenquelle")
    uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx"])
    st.info("Die App liest Blatt 1 als Servicemeldungen und Blatt 2 als Kosten/Stunden. Verbindung ueber Auftrag.")

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
    st.info("Kein zweites Tabellenblatt fuer Kosten/Stunden gefunden. Kostenbereiche bleiben leer.")
elif kosten_error:
    st.warning(kosten_error)

relevant_df = prepared_df[
    (~prepared_df["_Ausgeschlossen"]) & (~prepared_df["_Loeschvermerk"])
].copy()

# ------------------------------------------------------------
# Filter
# ------------------------------------------------------------

with st.sidebar:
    st.header("Filter")

    kategorien = sorted(relevant_df["_Kategorie"].dropna().unique().tolist())
    selected_kategorien = st.multiselect("Meldungskategorie", kategorien, default=kategorien)

    statuswerte = sorted(relevant_df["_Statusgruppe"].dropna().unique().tolist())
    selected_status = st.multiselect("Statusgruppe", statuswerte, default=statuswerte)

    selected_auftragsarten = []
    if "Auftragsart_Kostenblatt" in relevant_df.columns:
        auftragsarten = sorted(
            relevant_df["Auftragsart_Kostenblatt"].dropna().astype(str).replace("", pd.NA).dropna().unique().tolist()
        )
        selected_auftragsarten = st.multiselect("Auftragsart Kostenblatt", auftragsarten, default=auftragsarten)

    selected_tech = []
    if columns.get("techplatz"):
        tech_values = sorted(relevant_df[columns["techplatz"]].dropna().astype(str).unique().tolist())
        selected_tech = st.multiselect("Technischer Platz", tech_values, default=[])

    selected_prio = []
    if columns.get("prio"):
        prio_values = sorted(relevant_df[columns["prio"]].dropna().astype(str).unique().tolist())
        selected_prio = st.multiselect("Prioritaet", prio_values, default=[])

filtered_df = relevant_df[
    relevant_df["_Kategorie"].isin(selected_kategorien)
    & relevant_df["_Statusgruppe"].isin(selected_status)
].copy()

if selected_auftragsarten and "Auftragsart_Kostenblatt" in filtered_df.columns:
    filtered_df = filtered_df[filtered_df["Auftragsart_Kostenblatt"].astype(str).isin(selected_auftragsarten)]

if selected_tech and columns.get("techplatz"):
    filtered_df = filtered_df[filtered_df[columns["techplatz"]].astype(str).isin(selected_tech)]

if selected_prio and columns.get("prio"):
    filtered_df = filtered_df[filtered_df[columns["prio"]].astype(str).isin(selected_prio)]

# ------------------------------------------------------------
# KPIs
# ------------------------------------------------------------

offene_mf = int(filtered_df[(filtered_df["_Kategorie"] == "MF Reparatur") & filtered_df["_Offen_Status"]].shape[0])
critical_mf = int(filtered_df[filtered_df["_MF_groesser_10_Tage_offen"]].shape[0])
overdue_maintenance = int(filtered_df[filtered_df["_Wartung_im_Verzug"]].shape[0])
planned_maintenance = int(filtered_df[filtered_df["_Wartung_geplant"]].shape[0])
steuerungsrelevant_offen = int(filtered_df[filtered_df["_Steuerungsrelevant_offen"]].shape[0])
langlaeufer_30 = int(filtered_df[filtered_df["_Groesser_30_Tage_offen"]].shape[0])

if columns.get("techplatz"):
    fahrzeuge_in_werkstatt = int(
        filtered_df[filtered_df["_In_Werkstatt"]][columns["techplatz"]].dropna().astype(str).nunique()
    )
else:
    fahrzeuge_in_werkstatt = 0

gesamtkosten = float(filtered_df["Kosten_Ist"].sum())
gesamtstunden = float(filtered_df["Kosten_Stunden"].sum())
standortwert = float(filtered_df["Standortwert"].sum())
interne_stunden = float(filtered_df["Interne_Stunden"].sum())
interne_kosten = float(filtered_df["Interne_Kosten"].sum())
vlad_stunden = float(filtered_df["VLAD_Stunden"].sum())
vlad_erloese = float(filtered_df["VLAD_Erloese"].sum())
vlad_anteil = (vlad_stunden / gesamtstunden * 100) if gesamtstunden > 0 else 0.0

verrechenbare_auftraege = int(
    filtered_df[filtered_df["Verrechenbar"]]["_Auftrag_key"].replace("", pd.NA).dropna().nunique()
)
interne_auftraege = int(
    filtered_df[~filtered_df["Verrechenbar"]]["_Auftrag_key"].replace("", pd.NA).dropna().nunique()
)
deckungsbeitrag_sicht = vlad_erloese - interne_kosten

k1, k2, k3, k4 = st.columns(4)
k1.metric("Offene MF-Reparaturen", offene_mf)
k2.metric("MF >10 Tage offen", critical_mf)
k3.metric("Wartungen im Verzug", overdue_maintenance)
k4.metric("Aktuell in Werkstatt", fahrzeuge_in_werkstatt)

k5, k6, k7, k8 = st.columns(4)
k5.metric("Geplante Wartungen", planned_maintenance)
k6.metric("Steuerungsrelevant offen", steuerungsrelevant_offen)
k7.metric(">30 Tage offen", langlaeufer_30)
k8.metric("Istkosten gesamt", euro(gesamtkosten))

k9, k10, k11, k12 = st.columns(4)
k9.metric("Interne Stunden", hours(interne_stunden))
k10.metric("VLAD Stunden", hours(vlad_stunden))
k11.metric("VLAD Erloese", euro(vlad_erloese))
k12.metric("VLAD Anteil", percent(vlad_anteil))

st.divider()

# ------------------------------------------------------------
# Tabs
# ------------------------------------------------------------

tab0, tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "Werkstattleiter",
    "Uebersicht",
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
    st.write("Fokus auf steuerungsrelevante offene Arbeiten, Werkstattauslastung, Kosten und VLAD-Erloese.")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Kritische MF-Reparaturen")
        critical = filtered_df[filtered_df["_MF_groesser_10_Tage_offen"]].copy().sort_values("_Tage_offen", ascending=False)
        if critical.empty:
            st.success("Keine MF-Reparaturen ueber 10 Tage offen.")
        else:
            show_table(critical, columns, max_rows=20)

    with c2:
        st.markdown("### Wartungen im Verzug")
        overdue = filtered_df[filtered_df["_Wartung_im_Verzug"]].copy()
        if not overdue.empty:
            overdue["_Tage_im_Verzug"] = (pd.Timestamp(date.today()) - overdue["_Faelligkeit"]).dt.days
            overdue = overdue.sort_values("_Tage_im_Verzug", ascending=False)
            cols = [c for c in detail_columns(columns) + ["_Tage_im_Verzug"] if c in overdue.columns]
            display_dataframe(overdue[cols].head(20), use_container_width=True, hide_index=True)
        else:
            st.success("Keine ueberfaelligen Wartungen gefunden.")

    st.markdown("### Auslastung intern vs. VLAD")
    split_df = pd.DataFrame({
        "Bereich": ["Intern", "VLAD Leistung an Dritte"],
        "Stunden": [interne_stunden, vlad_stunden],
        "Wert": [interne_kosten, vlad_erloese],
    })
    c3, c4 = st.columns(2)
    with c3:
        fig = px.pie(split_df, names="Bereich", values="Stunden", title="Werkstattstunden intern vs. VLAD")
        st.plotly_chart(fig, use_container_width=True)
    with c4:
        fig = px.bar(split_df, x="Bereich", y="Wert", text="Wert", title="Interne Kosten vs. VLAD-Erloese")
        fig.update_traces(texttemplate="%{text:,.2f}", textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

with tab1:
    st.subheader("Uebersicht")
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
    st.subheader("Kritische MF-Reparaturen, laenger als 10 Tage offen")
    critical = filtered_df[filtered_df["_MF_groesser_10_Tage_offen"]].copy().sort_values("_Tage_offen", ascending=False)
    st.metric("Kritische MF-Reparaturen", int(critical.shape[0]))
    if critical.empty:
        st.success("Keine offenen MF-Reparaturen ueber 10 Tage gefunden.")
    else:
        show_table(critical, columns)
        make_download(critical, "kritische_mf_reparaturen.csv")

with tab3:
    st.subheader("Wartungen")
    st.write("Wartungen mit offenem Status und gewuenschtem Ende in der Zukunft werden als geplant dargestellt und nicht als Verzug gezaehlt.")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### Wartungen im Verzug")
        overdue = filtered_df[filtered_df["_Wartung_im_Verzug"]].copy()
        if not overdue.empty:
            overdue["_Tage_im_Verzug"] = (pd.Timestamp(date.today()) - overdue["_Faelligkeit"]).dt.days
            overdue = overdue.sort_values("_Tage_im_Verzug", ascending=False)
            cols = [c for c in detail_columns(columns) + ["_Tage_im_Verzug"] if c in overdue.columns]
            display_dataframe(overdue[cols], use_container_width=True, hide_index=True)
            make_download(overdue, "wartungen_im_verzug.csv")
        else:
            st.success("Keine ueberfaelligen Wartungen gefunden.")
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
            fig = px.bar(top_all, x="Anzahl", y=tech_col, orientation="h", title="Top 20 technische Plaetze nach Meldungen")
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            mf_df = filtered_df[filtered_df["_Kategorie"] == "MF Reparatur"]
            top_mf = mf_df.groupby(tech_col).size().reset_index(name="Anzahl_MF").sort_values("Anzahl_MF", ascending=False).head(20)
            fig = px.bar(top_mf, x="Anzahl_MF", y=tech_col, orientation="h", title="Top 20 technische Plaetze nach MF-Reparaturen")
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("### Technische Plaetze nach Stunden und Kosten")
        by_tech = filtered_df.groupby(tech_col).agg(
            Meldungen=("_Meldungsart", "size"),
            Auftraege=("_Auftrag_key", "nunique"),
            Iststunden=("Kosten_Stunden", "sum"),
            Istkosten=("Kosten_Ist", "sum"),
            Interne_Stunden=("Interne_Stunden", "sum"),
            Interne_Kosten=("Interne_Kosten", "sum"),
            VLAD_Stunden=("VLAD_Stunden", "sum"),
            VLAD_Erloese=("VLAD_Erloese", "sum"),
        ).reset_index().sort_values("Iststunden", ascending=False)
        display_dataframe(by_tech.head(100), use_container_width=True, hide_index=True)

with tab5:
    st.subheader("Standzeiten und Langlaeufer")
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
    st.write("Kosten und Stunden werden aus dem zweiten Tabellenblatt ueber die Auftragsnummer verbunden. VLAD wird als Leistung an Dritte getrennt ausgewiesen.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Istkosten gesamt", euro(gesamtkosten))
    c2.metric("Iststunden gesamt", hours(gesamtstunden))
    c3.metric("Interne Kosten", euro(interne_kosten))
    c4.metric("Interne Stunden", hours(interne_stunden))

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("VLAD Erloese", euro(vlad_erloese))
    c6.metric("VLAD Stunden", hours(vlad_stunden))
    c7.metric("VLAD Anteil Stunden", percent(vlad_anteil))
    c8.metric("VLAD Auftraege", verrechenbare_auftraege)

    c9, c10, c11, c12 = st.columns(4)
    c9.metric("Interne Auftraege", interne_auftraege)
    c10.metric("Plankosten", euro(filtered_df["Kosten_Plan"].sum()))
    c11.metric("Standortwert / Verrechnung", euro(standortwert))
    c12.metric("Erloese minus interne Kosten", euro(deckungsbeitrag_sicht))

    st.markdown("### Aufteilung intern vs. VLAD")
    split_hours = pd.DataFrame({"Bereich": ["Intern", "VLAD Leistung an Dritte"], "Stunden": [interne_stunden, vlad_stunden]})
    split_value = pd.DataFrame({"Bereich": ["Interne Kosten", "VLAD Erloese"], "Wert": [interne_kosten, vlad_erloese]})
    c13, c14 = st.columns(2)
    with c13:
        fig = px.pie(split_hours, names="Bereich", values="Stunden", title="Stundenaufteilung intern vs. VLAD")
        st.plotly_chart(fig, use_container_width=True)
    with c14:
        fig = px.bar(split_value, x="Bereich", y="Wert", text="Wert", title="Interne Kosten vs. VLAD-Erloese")
        fig.update_traces(texttemplate="%{text:,.2f}", textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Auftragsarten nach Stunden und Wert")
    by_art = filtered_df.groupby("Auftragsart_Kostenblatt").agg(
        Auftraege=("_Auftrag_key", "nunique"),
        Meldungen=("_Meldungsart", "size"),
        Iststunden=("Kosten_Stunden", "sum"),
        Istkosten=("Kosten_Ist", "sum"),
        Interne_Stunden=("Interne_Stunden", "sum"),
        Interne_Kosten=("Interne_Kosten", "sum"),
        VLAD_Stunden=("VLAD_Stunden", "sum"),
        VLAD_Erloese=("VLAD_Erloese", "sum"),
    ).reset_index().sort_values("Iststunden", ascending=False)
    display_dataframe(by_art, use_container_width=True, hide_index=True)

with tab7:
    st.subheader("Kostentreiber und Erloesbringer")
    if not columns.get("techplatz"):
        st.warning("Techn. Platz wurde nicht erkannt.")
    else:
        tech_col = columns["techplatz"]
        by_tech = filtered_df.groupby(tech_col).agg(
            Istkosten=("Kosten_Ist", "sum"),
            Iststunden=("Kosten_Stunden", "sum"),
            Plankosten=("Kosten_Plan", "sum"),
            Interne_Stunden=("Interne_Stunden", "sum"),
            Interne_Kosten=("Interne_Kosten", "sum"),
            VLAD_Stunden=("VLAD_Stunden", "sum"),
            VLAD_Erloese=("VLAD_Erloese", "sum"),
            Standortwert=("Standortwert", "sum"),
            Auftraege=("_Auftrag_key", "nunique"),
            Meldungen=("_Meldungsart", "size"),
        ).reset_index()

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("### Top 20 technische Plaetze nach Istkosten")
            top_cost = by_tech.sort_values("Istkosten", ascending=False).head(20)
            fig = px.bar(top_cost, x="Istkosten", y=tech_col, orientation="h", title="Top 20 technische Plaetze nach Istkosten")
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.markdown("### Top 20 technische Plaetze nach VLAD-Erloesen")
            top_vlad = by_tech.sort_values("VLAD_Erloese", ascending=False).head(20)
            fig = px.bar(top_vlad, x="VLAD_Erloese", y=tech_col, orientation="h", title="Top 20 technische Plaetze nach VLAD-Erloesen")
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("### Kostentreiber-Tabelle")
        by_tech = by_tech.sort_values("Istkosten", ascending=False)
        display_dataframe(by_tech, use_container_width=True, hide_index=True)
        make_download(by_tech, "kostentreiber_technischer_platz.csv")

    st.markdown("### Top 20 Auftraege nach VLAD-Erloesen")
    top_vlad_orders = filtered_df[filtered_df["Verrechenbar"]].copy().sort_values("VLAD_Erloese", ascending=False)
    if top_vlad_orders.empty:
        st.info("Keine VLAD-Auftraege im aktuellen Filter gefunden.")
    else:
        show_table(top_vlad_orders, columns, max_rows=20)
        make_download(top_vlad_orders, "top_vlad_erloese.csv")

    st.markdown("### Top 20 interne Kostentreiber")
    top_internal = filtered_df[~filtered_df["Verrechenbar"]].copy().sort_values("Interne_Kosten", ascending=False)
    if top_internal.empty:
        st.info("Keine internen Kostentreiber im aktuellen Filter gefunden.")
    else:
        show_table(top_internal, columns, max_rows=20)
        make_download(top_internal, "top_interne_kostentreiber.csv")

with tab8:
    st.subheader("Detaildaten")
    show_table(filtered_df, columns)
    make_download(filtered_df, "vat_kfz_servicemeldungen_gefiltert.csv")

# ------------------------------------------------------------
# Diagnosebereich
# ------------------------------------------------------------

with st.expander("Erkannte Spalten / Diagnose"):
    st.write("Tabellenblaetter:")
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
        display_dataframe(status_diag, use_container_width=True, hide_index=True)
    st.write("Kosten-Merge Diagnose:")
    st.write({
        "Meldungen_gesamt": int(prepared_df.shape[0]),
        "Meldungen_mit_Auftrag": int((prepared_df["_Auftrag_key"] != "").sum()),
        "Meldungen_mit_Kostenzeile": int((prepared_df["Kostenzeilen"] > 0).sum()),
        "Summe_Istkosten": float(prepared_df["Kosten_Ist"].sum()),
        "Summe_Iststunden": float(prepared_df["Kosten_Stunden"].sum()),
        "Summe_Interne_Kosten": float(prepared_df["Interne_Kosten"].sum()),
        "Summe_Interne_Stunden": float(prepared_df["Interne_Stunden"].sum()),
        "Summe_VLAD_Erloese": float(prepared_df["VLAD_Erloese"].sum()),
        "Summe_VLAD_Stunden": float(prepared_df["VLAD_Stunden"].sum()),
        "Summe_Standortwert": float(prepared_df["Standortwert"].sum()),
    })
