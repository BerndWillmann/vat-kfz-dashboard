import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import date

# ============================================================
# VAT KFZ Werkstätten-Cockpit V2.4
# SAP-Servicemeldungen + Kosten/Stunden aus zweitem Tabellenblatt
# Neu: Trennung intern / VLAD
# VLAD = Leistung an Dritte = externe, verrechenbare Leistung / Erlöse
# ============================================================

st.set_page_config(
    page_title="VAT KFZ Werkstätten-Cockpit",
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


# ------------------------------------------------------------
# Allgemeine Hilfsfunktionen
# ------------------------------------------------------------

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
    """Sucht eine Spalte robust anhand möglicher Namen."""
    if df is None:
        return None

    columns = list(df.columns)
    normalized = {clean_colname(c).lower(): c for c in columns}

    for candidate in candidates:
        candidate_lower = candidate.lower().strip()
        if candidate_lower in normalized:
            return normalized[candidate_lower]

    if contains:
        for candidate in candidates:
            candidate_lower = candidate.lower().strip()
            for col in columns:
                col_lower = clean_colname(col).lower()
                if candidate_lower in col_lower:
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

    # Leerzeichen entfernen
    s = s.str.replace(" ", "", regex=False)

    # leere/ungültige Werte
    s = s.replace(["", "nan", "NaN", "None", "-", "–"], "0")

    # Wenn Dezimalkomma vorhanden ist: Tausenderpunkt entfernen und Komma zu Punkt
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
    """Zahlenformat AT/DE: Tausenderpunkt und Dezimalkomma."""
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
    """Formatiert Tabellen für die Anzeige, ohne die Originaldaten zu verändern."""
    if df is None:
        return df

    display_df = df.copy()

    euro_cols = [
        "Kosten_Ist",
        "Kosten_Plan",
        "Standortwert",
        "Interne_Kosten",
        "VLAD_Erloese",
        "Erlöse",
        "Istkosten",
        "Plankosten",
        "Materialkosten",
    ]

    hour_cols = [
        "Kosten_Stunden",
        "Interne_Stunden",
        "VLAD_Stunden",
        "Iststunden",
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
        "prio": find_column(df, ["Priorität", "Prio"], contains=True),
        "meldedatum": find_column(df, ["Meldedatum", "Meld.datum", "Erfassungsdatum"], contains=True),
        "faellig": find_column(df, ["Gewünschtes Ende", "Gew. Ende", "Fälligkeit", "Faelligkeit"], contains=True),
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

    if columns["meldedatum"]:
        df["_Meldungsdatum"] = safe_to_datetime(df[columns["meldedatum"]])
    else:
        df["_Meldungsdatum"] = pd.NaT

    if columns["faellig"]:
        df["_Faelligkeit"] = safe_to_datetime(df[columns["faellig"]])
    else:
        df["_Faelligkeit"] = pd.NaT

    if columns["progn_ende"]:
        df["_Prognose_Ende"] = safe_to_datetime(df[columns["progn_ende"]])
    else:
        df["_Prognose_Ende"] = pd.NaT

    if columns["abschluss"]:
        df["_Abschlussdatum"] = safe_to_datetime(df[columns["abschluss"]])
    else:
        df["_Abschlussdatum"] = pd.NaT

    today = pd.Timestamp(date.today())

    df["_Tage_offen"] = (today - df["_Meldungsdatum"]).dt.days
    df.loc[df["_Abgeschlossen_Status"], "_Tage_offen"] = pd.NA

    df["_Dauer_bis_Abschluss"] = (df["_Abschlussdatum"] - df["_Meldungsdatum"]).dt.days

    df["_Ampel_Standzeit"] = df["_Tage_offen"].apply(status_ampel)

    df["_MF_groesser_10_Tage_offen"] = (
        (df["_Kategorie"] == "MF Reparatur")
        & (df["_Offen_Status"])
        & (df["_Tage_offen"].fillna(0) > KRITISCHE_REPARATUR_TAGE)
    )

    df["_Groesser_30_Tage_offen"] = (
        (df["_Offen_Status"])
        & (df["_Tage_offen"].fillna(0) > LANG_LAEUFER_TAGE)
    )

    # Wartungslogik:
    # Eine Wartung ist im Verzug, wenn sie offen ist und das gewünschte Ende in der Vergangenheit liegt.
    df["_Wartung_im_Verzug"] = (
        (df["_Kategorie"] == "WE/WK Wartung")
        & (df["_Offen_Status"])
        & (df["_Faelligkeit"].notna())
        & (df["_Faelligkeit"] < today)
    )

    # Eine Wartung ist geplant, wenn sie offen ist und das gewünschte Ende heute oder in der Zukunft liegt.
    df["_Wartung_geplant"] = (
        (df["_Kategorie"] == "WE/WK Wartung")
        & (df["_Offen_Status"])
        & (df["_Faelligkeit"].notna())
        & (df["_Faelligkeit"] >= today)
    )

    df["_Steuerungsrelevant_offen"] = (
        (df["_Offen_Status"])
        & (~df["_Ausgeschlossen"])
        & (~df["_Loeschvermerk"])
        & (df["_Kategorie"].isin(["MF Reparatur", "WE/WK Wartung", "G Garantie", "F Meldung", "Sonstige"]))
    )

    # Leere Werte für Kostenfelder vorbereiten
    for col in [
        "Kosten_Stunden",
        "Kosten_Ist",
        "Kosten_Plan",
        "Standortwert",
        "Kostenzeilen",
        "Interne_Stunden",
        "Interne_Kosten",
        "VLAD_Stunden",
        "VLAD_Erloese",
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
        "auftragsart": find_column(df, ["Auftragsart", "Auftragsart Text", "AufArt"], contains=True),
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

    if kosten_cols["istarbeit"]:
        df["Kosten_Stunden"] = safe_to_number(df[kosten_cols["istarbeit"]])
    else:
        df["Kosten_Stunden"] = 0.0

    if kosten_cols["kosten_ist"]:
        df["Kosten_Ist"] = safe_to_number(df[kosten_cols["kosten_ist"]])
    else:
        df["Kosten_Ist"] = 0.0

    if kosten_cols["kosten_plan"]:
        df["Kosten_Plan"] = safe_to_number(df[kosten_cols["kosten_plan"]])
    else:
        df["Kosten_Plan"] = 0.0

    if kosten_cols["standortwert"]:
        df["Standortwert"] = safe_to_number(df[kosten_cols["standortwert"]])
    else:
        df["Standortwert"] = 0.0

    df["Verrechenbar"] = df["Auftragsart_Kostenblatt"].eq("VLAD")

    df["Interne_Stunden"] = df.apply(
        lambda r: 0.0 if r["Auftragsart_Kostenblatt"] == "VLAD" else r["Kosten_Stunden"],
        axis=1
    )

    df["Interne_Kosten"] = df.apply(
        lambda r: 0.0 if r["Auftragsart_Kostenblatt"] == "VLAD" else r["Kosten_Ist"],
        axis=1
    )

    df["VLAD_Stunden"] = df.apply(
        lambda r: r["Kosten_Stunden"] if r["Auftragsart_Kostenblatt"] == "VLAD" else 0.0,
        axis=1
    )

    # Fachliche Logik:
    # VLAD = Leistung an Dritte.
    # Der Istkosten-Wert wird hier im Dashboard als externer Erlös ausgewiesen.
    df["VLAD_Erloese"] = df.apply(
        lambda r: r["Kosten_Ist"] if r["Auftragsart_Kostenblatt"] == "VLAD" else 0.0,
        axis=1
    )

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

    # Wenn pro Auftrag mehrere Auftragsarten vorkommen, sichtbar machen
    art_agg = df.groupby("_Auftrag_key")["Auftragsart_Kostenblatt"].apply(
        lambda x: ", ".join(sorted(set([v for v in x if str(v).strip() != ""])))
    ).reset_index()

    agg = agg.merge(art_agg, on="_Auftrag_key", how="left")

    return agg, kosten_cols, None


def merge_costs(meldungen_df, kosten_agg):
    df = meldungen_df.copy()

    if kosten_agg is None:
        for col in [
            "Kosten_Stunden",
            "Kosten_Ist",
            "Kosten_Plan",
            "Standortwert",
            "Interne_Stunden",
            "Interne_Kosten",
            "VLAD_Stunden",
            "VLAD_Erloese",
            "Kostenzeilen",
        ]:
            df[col] = 0.0

        df["Auftragsart_Kostenblatt"] = ""
        df["Verrechenbar"] = False

        return df

    df = df.merge(
        kosten_agg,
        on="_Auftrag_key",
        how="left",
        suffixes=("", "_kosten")
    )

    numeric_cols = [
        "Kosten_Stunden",
        "Kosten_Ist",
        "Kosten_Plan",
        "Standortwert",
        "Interne_Stunden",
        "Interne_Kosten",
        "VLAD_Stunden",
        "VLAD_Erloese",
        "Kostenzeilen",
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
        "meldung",
        "auftrag",
        "art",
        "techplatz",
        "equipment",
        "kurztext",
        "status",
        "anwenderstatus",
        "prio",
        "meldedatum",
        "faellig",
        "progn_ende",
        "abschluss",
        "arbeitsplatz",
        "standort",
    ]:
        col = columns.get(key)
        if col and col not in cols:
            cols.append(col)

    add_cols = [
        "_Kategorie",
        "_Statusgruppe",
        "_Tage_offen",
        "_Dauer_bis_Abschluss",
        "Auftragsart_Kostenblatt",
        "Kosten_Stunden",
        "Kosten_Ist",
        "Kosten_Plan",
        "Interne_Stunden",
        "Interne_Kosten",
        "VLAD_Stunden",
        "VLAD_Erloese",
        "Standortwert",
        "Verrechenbar",
        "_Ampel_Standzeit",
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
# Oberfläche
# ------------------------------------------------------------

st.title("🔧 VAT KFZ Werkstätten-Cockpit V2.4")
st.caption(
    "SAP-Servicemeldungen plus Kosten/Stunden je Auftrag. "
    "VLAD wird als Leistung an Dritte getrennt als externe Stunden und Erlöse ausgewiesen."
)

with st.sidebar:
    st.header("Datenquelle")
    uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx"])
    st.info(
        "Die App liest Blatt 1 als Servicemeldungen und Blatt 2 als Kosten/Stunden. "
        "Die Verbindung erfolgt über die Auftragsnummer."
    )

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

relevant_df = prepared_df[
    (~prepared_df["_Ausgeschlossen"])
    & (~prepared_df["_Loeschvermerk"])
].copy()


# ------------------------------------------------------------
# Filter
# ------------------------------------------------------------

with st.sidebar:
    st.header("Filter")

    kategorien = sorted(relevant_df["_Kategorie"].dropna().unique().tolist())
    selected_kategorien = st.multiselect(
        "Meldungskategorie",
        kategorien,
        default=kategorien
    )

    statuswerte = sorted(relevant_df["_Statusgruppe"].dropna().unique().tolist())
    selected_status = st.multiselect(
        "Statusgruppe",
        statuswerte,
        default=statuswerte
    )

    selected_auftragsarten = []
    if "Auftragsart_Kostenblatt" in relevant_df.columns:
        auftragsarten = sorted(
            relevant_df["Auftragsart_Kostenblatt"]
            .dropna()
            .astype(str)
            .replace("", pd.NA)
            .dropna()
            .unique()
            .tolist()
        )
        selected_auftragsarten = st.multiselect(
            "Auftragsart Kostenblatt",
            auftragsarten,
            default=auftragsarten
        )

    selected_tech = []
    if columns.get("techplatz"):
        tech_values = sorted(
            relevant_df[columns["techplatz"]]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )
        selected_tech = st.multiselect(
            "Technischer Platz",
            tech_values,
            default=[]
        )

    selected_prio = []
    if columns.get("prio"):
        prio_values = sorted(
            relevant_df[columns["prio"]]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )
        selected_prio = st.multiselect(
            "Priorität",
            prio_values,
            default=[]
        )

filtered_df = relevant_df[
    relevant_df["_Kategorie"].isin(selected_kategorien)
    & relevant_df["_Statusgruppe"].isin(selected_status)
].copy()

if selected_auftragsarten and "Auftragsart_Kostenblatt" in filtered_df.columns:
    filtered_df = filtered_df[
        filtered_df["Auftragsart_Kostenblatt"].astype(str).isin(selected_auftragsarten)
    ]

if selected_tech and columns.get("techplatz"):
    filtered_df = filtered_df[
        filtered_df[columns["techplatz"]].astype(str).isin(selected_tech)
    ]

if selected_prio and columns.get("prio"):
    filtered_df = filtered_df[
        filtered_df[columns["prio"]].astype(str).isin(selected_prio)
    ]


# ------------------------------------------------------------
# KPI-Berechnung
# ------------------------------------------------------------

offene_mf = int(
    filtered_df[
        (filtered_df["_Kategorie"] == "MF Reparatur")
        & (filtered_df["_Offen_Status"])
    ].shape[0]
)

critical_mf = int(filtered_df[filtered_df["_MF_groesser_10_Tage_offen"]].shape[0])
overdue_maintenance = int(filtered_df[filtered_df["_Wartung_im_Verzug"]].shape[0])
planned_maintenance = int(filtered_df[filtered_df["_Wartung_geplant"]].shape[0])
steuerungsrelevant_offen = int(filtered_df[filtered_df["_Steuerungsrelevant_offen"]].shape[0])
langlaeufer_30 = int(filtered_df[filtered_df["_Groesser_30_Tage_offen"]].shape[0])

if columns.get("techplatz"):
    fahrzeuge_in_werkstatt = int(
        filtered_df[filtered_df["_In_Werkstatt"]][columns["techplatz"]]
        .dropna()
        .astype(str)
        .nunique()
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

if gesamtstunden > 0:
    vlad_anteil = (vlad_stunden / gesamtstunden) * 100
else:
    vlad_anteil = 0.0

verrechenbare_auftraege = int(
    filtered_df[filtered_df["Verrechenbar"]]["_Auftrag_key"].replace("", pd.NA).dropna().nunique()
)

interne_auftraege = int(
    filtered_df[~filtered_df["Verrechenbar"]]["_Auftrag_key"].replace("", pd.NA).dropna().nunique()
)

deckungsbeitrag_sicht = vlad_erloese - interne_kosten


# ------------------------------------------------------------
# KPI-Anzeige
# ------------------------------------------------------------

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
k11.metric("VLAD Erlöse", euro(vlad_erloese))
k12.metric("VLAD Anteil", percent(vlad_anteil))

st.divider()


# ---------------------------------------------------
