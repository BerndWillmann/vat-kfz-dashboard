import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import date

# ============================================================
# VAT KFZ Werkstaetten-Cockpit V2.1
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
GARANTIE_PREFIXE = ["GM", "G"]
F_PREFIXE = ["F"]
KRITISCHE_REPARATUR_TAGE = 10
LANG_LAEUFER_TAGE = 30
WARTUNG_VORSCHAU_TAGE = 30

# SAP-Systemstatus nach Vorgabe VAT KFZ
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

OFFEN_STATUS = [
    "MOFN",
]

LOESCHVERMERK_STATUS = [
    "LÖVM MAUF MMAB",
    "LÖVM MMAB",
    "LOEVM MAUF MMAB",
    "LOEVM MMAB",
]

# -----------------------------
# Hilfsfunktionen
# -----------------------------
def clean_colname(col):
    return str(col).strip().replace("\n", " ").replace("  ", " ")


def normalize_status(value):
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().upper().split())


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

    # Fallback fuer SAP-Kombinationen
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


def load_excel(uploaded_file):
    if uploaded_file is not None:
        return pd.read_excel(uploaded_file, engine="openpyxl")

    default_file = "Servicemeldungen_VAT_KFZ.xlsx"
    try:
        return pd.read_excel(default_file, engine="openpyxl")
    except Exception:
        return None


def get_existing_column(df, name):
    return name if name in df.columns else None


def prepare_data(df):
    df = df.copy()
    df.columns = [clean_colname(c) for c in df.columns]

    # Exakte Spalten aus deinem SAP-Export werden bevorzugt.
    col_art = get_existing_column(df, "Meldungsart") or find_column(df, ["Meldungsart", "Meldungsart Text", "Art", "Meld.Art", "MArt"], contains=True)
    col_techplatz = get_existing_column(df, "Techn. Platz") or find_column(df, ["Technischer Platz", "Techn. Platz", "Techn Platz", "Equipment", "Equipement", "Objekt", "TechnPlatz"], contains=True)
    col_equipment = get_existing_column(df, "Equipment") or find_column(df, ["Equipment", "Equipement"], contains=True)
    col_meldung = get_existing_column(df, "Meldung") or find_column(df, ["Meldung", "Meldungsnummer", "Servicemeldung", "Ursprungsmeldung", "Meldungsnr"], contains=True)
    col_status = get_existing_column(df, "Systemstatus") or find_column(df, ["Systemstatus", "Status", "Anwenderstatus", "Folgestatus"], contains=True)
    col_anwenderstatus = get_existing_column(df, "Anwenderstat.") or find_column(df, ["Anwenderstat.", "Anwenderstatus"], contains=True)
    col_meldedatum = get_existing_column(df, "Angelegt am") or find_column(df, ["Meldungsdatum", "Angelegt am", "Erfassungsdatum", "Startdatum", "Meldedatum", "Datum"], contains=True)

    # Gew.Ende steuert Wartungen. Falls leer, wird progn. Ende als Ersatz verwendet.
    col_faellig = get_existing_column(df, "Gew.Ende") or get_existing_column(df, "progn. Ende") or find_column(df, ["Gew.Ende", "progn. Ende", "Fällig am", "Faellig am", "Fälligkeit", "Faelligkeit", "Endtermin", "Gewünschtes Ende", "Geplantes Ende", "Sollende"], contains=True)
    col_progn_ende = get_existing_column(df, "progn. Ende") or find_column(df, ["progn. Ende", "Prognostiziertes Ende"], contains=True)

    col_abschluss = get_existing_column(df, "Abschlußdatum") or get_existing_column(df, "Abschlussdatum") or find_column(df, ["Abschlußdatum", "Abschlussdatum", "Erledigt am", "Fertigstellung", "Rueckmeldedatum", "Rückmeldedatum", "Istende"], contains=True)

    col_prio = get_existing_column(df, "Priorität") or find_column(df, ["Priorität", "Prioritaet", "Prio"], contains=True)
    col_kurztext = get_existing_column(df, "Beschreibung") or find_column(df, ["Kurztext", "Beschreibung", "Meldungstext", "Text", "Schaden"], contains=True)
    col_auftrag = get_existing_column(df, "Auftrag") or find_column(df, ["Auftrag", "Auftragsnummer", "IH-Auftrag"], contains=True)
    col_langtext = get_existing_column(df, "Langtext i") or find_column(df, ["Langtext i", "Langtext", "Langtext intern"], contains=True)
    col_codegruppe = get_existing_column(df, "Codegruppe") or find_column(df, ["Codegruppe"], contains=True)
    col_codierung = get_existing_column(df, "Codierung") or find_column(df, ["Codierung"], contains=True)
    col_standort = get_existing_column(df, "Standort") or find_column(df, ["Standort"], contains=True)
    col_arbeitsplatz = get_existing_column(df, "Verantw.ArbPl.") or find_column(df, ["Verantw.ArbPl.", "Verantw. ArbPl.", "Arbeitsplatz"], contains=True)

    required = {
        "Meldungsart": col_art,
        "Technischer Platz / Equipment": col_techplatz,
        "Meldungsdatum": col_meldedatum,
        "Systemstatus": col_status,
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

    if col_progn_ende:
        progn = safe_to_datetime(df[col_progn_ende])
        df["_Faelligkeit"] = df["_Faelligkeit"].fillna(progn)

    if col_abschluss:
        df["_Abschlussdatum"] = safe_to_datetime(df[col_abschluss])
    else:
        df["_Abschlussdatum"] = pd.NaT

    today = pd.Timestamp(date.today())

    if col_status:
        df["_Systemstatus_norm"] = df[col_status].apply(normalize_status)
        df["_Statusgruppe"] = df[col_status].apply(classify_status)
    else:
        df["_Systemstatus_norm"] = ""
        df["_Statusgruppe"] = "Unbekannt"

    df["_Loeschvermerk"] = df["_Statusgruppe"].eq("Loeschvermerk")
    df["_Abgeschlossen"] = df["_Statusgruppe"].eq("Abgeschlossen")
    df["_Offen_Status"] = df["_Statusgruppe"].isin(["Offen", "In Arbeit", "Unbekannt"])

    df["_Tage_offen"] = (today - df["_Meldungsdatum"]).dt.days
    df.loc[df["_Tage_offen"] < 0, "_Tage_offen"] = pd.NA

    df["_Dauer_bis_Abschluss"] = (df["_Abschlussdatum"] - df["_Meldungsdatum"]).dt.days
    df.loc[df["_Dauer_bis_Abschluss"] < 0, "_Dauer_bis_Abschluss"] = pd.NA

    # ------------------------------------------------------------
    # Wartungslogik V2.1
    # Eine Wartung mit offenem Status gilt NICHT als offene Wartung,
    # solange Gew.Ende noch nicht erreicht ist.
    # ------------------------------------------------------------
    df["_Wartung_geplant"] = (
        df["_Kategorie"].eq("WE/WK Wartung")
        & df["_Offen_Status"]
        & df["_Faelligkeit"].notna()
        & (df["_Faelligkeit"] >= today)
    )

    df["_Wartung_naechste_30_Tage"] = (
        df["_Wartung_geplant"]
        & (df["_Faelligkeit"] <= today + pd.Timedelta(days=WARTUNG_VORSCHAU_TAGE))
    )

    df["_Wartung_im_Verzug"] = (
        df["_Kategorie"].eq("WE/WK Wartung")
        & df["_Offen_Status"]
        & df["_Faelligkeit"].notna()
        & (df["_Faelligkeit"] < today)
    )

    df["_Wartung_ohne_Termin_offen"] = (
        df["_Kategorie"].eq("WE/WK Wartung")
        & df["_Offen_Status"]
        & df["_Faelligkeit"].isna()
    )

    # Steuerungsrelevant offen:
    # - Reparaturen, Garantie, F und Sonstige zählen bei offenem Status.
    # - Wartungen zählen nur dann, wenn sie im Verzug sind oder keinen Termin haben.
    df["_Steuerungsrelevant_offen"] = (
        (df["_Kategorie"].ne("WE/WK Wartung") & df["_Offen_Status"])
        | df["_Wartung_im_Verzug"]
        | df["_Wartung_ohne_Termin_offen"]
    )

    df["_In_Werkstatt"] = (
        df["_Statusgruppe"].isin(["Offen", "In Arbeit"])
        & (
            df["_Kategorie"].ne("WE/WK Wartung")
            | df["_Wartung_im_Verzug"]
            | df["_Wartung_ohne_Termin_offen"]
        )
    )

    df["_MF_groesser_10_Tage_offen"] = (
        df["_Kategorie"].eq("MF Reparatur")
        & df["_Offen_Status"]
        & df["_Tage_offen"].notna()
        & (df["_Tage_offen"] > KRITISCHE_REPARATUR_TAGE)
    )

    df["_Groesser_30_Tage_offen"] = (
        df["_Steuerungsrelevant_offen"]
        & df["_Tage_offen"].notna()
        & (df["_Tage_offen"] > LANG_LAEUFER_TAGE)
    )

    df["_Ampel_Standzeit"] = df["_Tage_offen"].apply(status_ampel)

    columns = {
        "art": col_art,
        "techplatz": col_techplatz,
        "equipment": col_equipment,
        "meldung": col_meldung,
        "status": col_status,
        "anwenderstatus": col_anwenderstatus,
        "meldedatum": col_meldedatum,
        "faellig": col_faellig,
        "progn_ende": col_progn_ende,
        "abschluss": col_abschluss,
        "prio": col_prio,
        "kurztext": col_kurztext,
        "auftrag": col_auftrag,
        "langtext": col_langtext,
        "codegruppe": col_codegruppe,
        "codierung": col_codierung,
        "standort": col_standort,
        "arbeitsplatz": col_arbeitsplatz,
    }
    return df, columns, missing_required


def detail_columns(columns):
    cols = []
    for key in [
        "meldung", "auftrag", "art", "techplatz", "equipment", "kurztext", "status",
        "anwenderstatus", "prio", "meldedatum", "faellig", "progn_ende", "abschluss",
        "arbeitsplatz", "standort"
    ]:
        col = columns.get(key)
        if col and col not in cols:
            cols.append(col)

    for col in [
        "_Kategorie", "_Statusgruppe", "_Tage_offen", "_Dauer_bis_Abschluss", "_Ampel_Standzeit"
    ]:
        if col not in cols:
            cols.append(col)
    return [c for c in cols if c is not None]


def make_download(df, filename):
    csv = df.to_csv(index=False, sep=";").encode("utf-8-sig")
    st.download_button(
        label=f"{filename} herunterladen",
        data=csv,
        file_name=filename,
        mime="text/csv",
    )


def show_table(df, columns, max_rows=None):
    cols = [c for c in detail_columns(columns) if c in df.columns]
    if max_rows:
        st.dataframe(df[cols].head(max_rows), use_container_width=True, hide_index=True)
    else:
        st.dataframe(df[cols], use_container_width=True, hide_index=True)


# -----------------------------
# Oberfläche
# -----------------------------
st.title("🔧 VAT KFZ Werkstätten-Cockpit V2.1")
st.caption("SAP-Servicemeldungen: MF-Reparaturen, Wartungstermine, Verzug, technische Plätze, Standzeiten und Werkstattauslastung")

with st.sidebar:
    st.header("Datenquelle")
    uploaded_file = st.file_uploader("Excel-Datei hochladen", type=["xlsx"])
    st.info("Du kannst die Datei hochladen oder als 'Servicemeldungen_VAT_KFZ.xlsx' direkt ins GitHub-Repository legen.")

    st.header("Regeln")
    st.write("**MM** wird ausgeschlossen")
    st.write("**LÖVM** wird ausgeschlossen")
    st.write("**MF** = wichtige Reparaturen")
    st.write("**WE/WK/W1/W2/W3/WZ/WP** = Wartungen")
    st.write("**Wartungen mit offenem Status zählen erst als offen, wenn Gew.Ende erreicht oder überschritten ist**")
    st.write("**G/GM** = Garantie")
    st.write("**F** = eigene Kategorie")
    st.write(f"MF kritisch ab **>{KRITISCHE_REPARATUR_TAGE} Tagen offen**")
    st.write(f"Langläufer ab **>{LANG_LAEUFER_TAGE} Tagen offen**")

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

# MM und Loeschvermerke ausschliessen
relevant_df = prepared_df[
    (~prepared_df["_Ausgeschlossen"])
    & (~prepared_df["_Loeschvermerk"])
].copy()

# -----------------------------
# Sidebar-Filter
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

filtered_df = relevant_df[
    relevant_df["_Kategorie"].isin(selected_kategorien)
    & relevant_df["_Statusgruppe"].isin(selected_status)
].copy()

if selected_tech and columns.get("techplatz"):
    filtered_df = filtered_df[filtered_df[columns["techplatz"]].astype(str).isin(selected_tech)]

if selected_prio and columns.get("prio"):
    filtered_df = filtered_df[filtered_df[columns["prio"]].astype(str).isin(selected_prio)]

# -----------------------------
# KPI-Bereich
# -----------------------------
offene_mf = int(filtered_df[(filtered_df["_Kategorie"] == "MF Reparatur") & filtered_df["_Offen_Status"]].shape[0])
critical_mf = int(filtered_df[filtered_df["_MF_groesser_10_Tage_offen"]].shape[0])
overdue_maintenance = int(filtered_df[filtered_df["_Wartung_im_Verzug"]].shape[0])
planned_maintenance = int(filtered_df[filtered_df["_Wartung_geplant"]].shape[0])
wartung_30 = int(filtered_df[filtered_df["_Wartung_naechste_30_Tage"]].shape[0])
wartung_ohne_termin = int(filtered_df[filtered_df["_Wartung_ohne_Termin_offen"]].shape[0])
offen_status_total = int(filtered_df[filtered_df["_Offen_Status"]].shape[0])
steuerungsrelevant_offen = int(filtered_df[filtered_df["_Steuerungsrelevant_offen"]].shape[0])
gesamt_relevant = int(filtered_df.shape[0])
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

k1, k2, k3, k4 = st.columns(4)
k1.metric("Offene MF-Reparaturen", offene_mf)
k2.metric("MF >10 Tage offen", critical_mf)
k3.metric("Wartungen im Verzug", overdue_maintenance)
k4.metric("Aktuell in Werkstatt", fahrzeuge_in_werkstatt)

k5, k6, k7, k8 = st.columns(4)
k5.metric("Geplante Wartungen", planned_maintenance)
k6.metric("Wartungen nächste 30 Tage", wartung_30)
k7.metric(">30 Tage steuerungsrelevant offen", langlaeufer_30)
k8.metric("Steuerungsrelevant offen", steuerungsrelevant_offen)

k9, k10, k11, k12 = st.columns(4)
k9.metric("Status offen gesamt", offen_status_total)
k10.metric("Wartungen ohne Termin", wartung_ohne_termin)
k11.metric("Relevante Meldungen ohne MM/LÖVM", gesamt_relevant)
k12.metric("Gewünschtes Ende Feld", columns.get("faellig") or "nicht erkannt")

st.divider()

# -----------------------------
# Tabs
# -----------------------------
tab0, tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "Werkstattleiter",
    "Übersicht",
    "Kritische MF-Reparaturen",
    "Wartungen",
    "Equipment-Analyse",
    "Standzeiten",
    "Detaildaten",
])

with tab0:
    st.subheader("Werkstattleiter-Ansicht")
    st.write("Fokus auf steuerungsrelevante offene Arbeiten. Geplante Wartungen mit Gew.Ende in der Zukunft werden separat angezeigt und nicht als offene Wartung gewertet.")

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("### Kritische MF-Reparaturen")
        critical = filtered_df[filtered_df["_MF_groesser_10_Tage_offen"]].copy()
        critical = critical.sort_values("_Tage_offen", ascending=False)
        if critical.empty:
            st.success("Keine MF-Reparaturen über 10 Tage offen.")
        else:
            show_table(critical, columns, max_rows=20)

    with c2:
        st.markdown("### Wartungen im Verzug")
        overdue = filtered_df[filtered_df["_Wartung_im_Verzug"]].copy()
        overdue["_Tage_im_Verzug"] = (pd.Timestamp(date.today()) - overdue["_Faelligkeit"]).dt.days
        overdue = overdue.sort_values("_Tage_im_Verzug", ascending=False)
        if overdue.empty:
            st.success("Keine überfälligen Wartungen gefunden.")
        else:
            cols = [c for c in detail_columns(columns) + ["_Tage_im_Verzug"] if c in overdue.columns]
            st.dataframe(overdue[cols].head(20), use_container_width=True, hide_index=True)

    if columns.get("techplatz"):
        st.markdown("### Top-Störfahrzeuge / Technische Plätze")
        top_mf_open = filtered_df[(filtered_df["_Kategorie"] == "MF Reparatur") & filtered_df["_Offen_Status"]]
        top_mf_open = top_mf_open.groupby(columns["techplatz"]).size().reset_index(name="Offene_MF")
        top_mf_open = top_mf_open.sort_values("Offene_MF", ascending=False).head(20)
        if not top_mf_open.empty:
            fig = px.bar(top_mf_open, x="Offene_MF", y=columns["techplatz"], orientation="h", title="Top 20 technische Plätze nach offenen MF-Reparaturen")
            fig.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Keine offenen MF-Reparaturen für die Top-Störer-Auswertung vorhanden.")

with tab1:
    st.subheader("Übersicht nach Meldungsart und Status")
    c1, c2 = st.columns(2)

    with c1:
        if filtered_df.empty:
            st.info("Keine Daten nach Filterauswahl vorhanden.")
        else:
            cat_count = filtered_df.groupby("_Kategorie").size().reset_index(name="Anzahl")
            fig = px.bar(cat_count, x="_Kategorie", y="Anzahl", title="Anzahl Meldungen nach Kategorie", text="Anzahl")
            st.plotly_chart(fig, use_container_width=True)

    with c2:
        if filtered_df.empty:
            st.info("Keine Daten nach Filterauswahl vorhanden.")
        else:
            status_count = filtered_df.groupby(["_Kategorie", "_Statusgruppe"]).size().reset_index(name="Anzahl")
            fig = px.bar(status_count, x="_Kategorie", y="Anzahl", color="_Statusgruppe", title="Status nach Kategorie", text="Anzahl")
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
        show_table(critical, columns)
        make_download(critical, "kritische_mf_reparaturen.csv")

with tab3:
    st.subheader("Wartungen")
    st.write("Wartungen mit offenem Status und Gew.Ende in der Zukunft werden als geplant dargestellt. Sie zählen nicht als offene Wartung im Verzug.")

    c1, c2 = st.columns(2)

    with c1:
        st.markdown("### Wartungen im Verzug")
        overdue = filtered_df[filtered_df["_Wartung_im_Verzug"]].copy()
        overdue["_Tage_im_Verzug"] = (pd.Timestamp(date.today()) - overdue["_Faelligkeit"]).dt.days
        overdue = overdue.sort_values("_Tage_im_Verzug", ascending=False)
        st.metric("Wartungen im Verzug", int(overdue.shape[0]))
        if overdue.empty:
            st.success("Keine überfälligen Wartungen gefunden.")
        else:
            cols = [c for c in detail_columns(columns) + ["_Tage_im_Verzug"] if c in overdue.columns]
            st.dataframe(overdue[cols], use_container_width=True, hide_index=True)
            make_download(overdue, "wartungen_im_verzug.csv")

    with c2:
        st.markdown("### Wartungen nächste 30 Tage")
        upcoming = filtered_df[filtered_df["_Wartung_naechste_30_Tage"]].copy()
        upcoming = upcoming.sort_values("_Faelligkeit", ascending=True)
        st.metric("Wartungen nächste 30 Tage", int(upcoming.shape[0]))
        if upcoming.empty:
            st.info("Keine offenen Wartungen in den nächsten 30 Tagen gefunden.")
        else:
            show_table(upcoming, columns)
            make_download(upcoming, "wartungen_naechste_30_tage.csv")

    c3, c4 = st.columns(2)
    with c3:
        st.markdown("### Geplante Wartungen")
        planned = filtered_df[filtered_df["_Wartung_geplant"]].copy()
        planned = planned.sort_values("_Faelligkeit", ascending=True)
        st.metric("Geplante Wartungen", int(planned.shape[0]))
        if planned.empty:
            st.info("Keine geplanten offenen Wartungen mit zukünftigem Gew.Ende gefunden.")
        else:
            show_table(planned, columns, max_rows=100)
            make_download(planned, "geplante_wartungen.csv")

    with c4:
        st.markdown("### Wartungen ohne Termin")
        no_due = filtered_df[filtered_df["_Wartung_ohne_Termin_offen"]].copy()
        no_due = no_due.sort_values("_Tage_offen", ascending=False)
        st.metric("Wartungen ohne Termin", int(no_due.shape[0]))
        if no_due.empty:
            st.success("Keine offenen Wartungen ohne Gew.Ende/progn. Ende gefunden.")
        else:
            show_table(no_due, columns, max_rows=100)
            make_download(no_due, "wartungen_ohne_termin.csv")

with tab4:
    st.subheader("Equipment-Analyse / Technischer Platz")

    if not columns.get("techplatz"):
        st.warning("Die Spalte 'Technischer Platz' bzw. 'Techn. Platz' wurde nicht erkannt.")
    else:
        tech_col = columns["techplatz"]

        c1, c2 = st.columns(2)
        with c1:
            top_all = filtered_df.groupby(tech_col).size().reset_index(name="Anzahl")
            top_all = top_all.sort_values("Anzahl", ascending=False).head(20)
            if top_all.empty:
                st.info("Keine Daten für Equipment-Analyse vorhanden.")
            else:
                fig = px.bar(top_all, x="Anzahl", y=tech_col, orientation="h", title="Top 20 technische Plätze nach allen relevanten Meldungen")
                fig.update_layout(yaxis={"categoryorder": "total ascending"})
                st.plotly_chart(fig, use_container_width=True)

        with c2:
            mf_df = filtered_df[filtered_df["_Kategorie"] == "MF Reparatur"]
            top_mf = mf_df.groupby(tech_col).size().reset_index(name="Anzahl_MF")
            top_mf = top_mf.sort_values("Anzahl_MF", ascending=False).head(20)
            if top_mf.empty:
                st.info("Keine MF-Daten für Equipment-Analyse vorhanden.")
            else:
                fig = px.bar(top_mf, x="Anzahl_MF", y=tech_col, orientation="h", title="Top 20 technische Plätze nach MF-Reparaturen")
                fig.update_layout(yaxis={"categoryorder": "total ascending"})
                st.plotly_chart(fig, use_container_width=True)

        st.subheader("Technische Plätze mit steuerungsrelevanten offenen Meldungen")
        open_by_equipment = filtered_df[filtered_df["_Steuerungsrelevant_offen"]].groupby(tech_col).agg(
            Steuerungsrelevant_offen=("_Steuerungsrelevant_offen", "size"),
            MF_offen=("_Kategorie", lambda x: (x == "MF Reparatur").sum()),
            Wartungen_im_Verzug=("_Wartung_im_Verzug", "sum"),
            Max_Tage_offen=("_Tage_offen", "max"),
            Durchschnitt_Tage_offen=("_Tage_offen", "mean"),
        ).reset_index().sort_values(["MF_offen", "Wartungen_im_Verzug", "Max_Tage_offen"], ascending=False)

        if open_by_equipment.empty:
            st.success("Keine steuerungsrelevanten offenen Meldungen je technischem Platz vorhanden.")
        else:
            st.dataframe(open_by_equipment, use_container_width=True, hide_index=True)
            make_download(open_by_equipment, "technische_plaetze_steuerungsrelevant_offen.csv")

with tab5:
    st.subheader("Standzeiten und Langläufer")
    st.write("Geplante Wartungen mit Gew.Ende in der Zukunft werden hier nicht als Langläufer gezählt.")
    stand = filtered_df[filtered_df["_Steuerungsrelevant_offen"]].copy()
    stand = stand.sort_values("_Tage_offen", ascending=False)

    if stand.empty:
        st.success("Keine steuerungsrelevanten offenen Meldungen gefunden.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            ampel = stand.groupby("_Ampel_Standzeit").size().reset_index(name="Anzahl")
            fig = px.pie(ampel, names="_Ampel_Standzeit", values="Anzahl", title="Standzeit-Ampel steuerungsrelevanter offener Meldungen")
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            stand_cat = stand.groupby(["_Kategorie", "_Ampel_Standzeit"]).size().reset_index(name="Anzahl")
            fig = px.bar(stand_cat, x="_Kategorie", y="Anzahl", color="_Ampel_Standzeit", title="Standzeit nach Kategorie", text="Anzahl")
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("Längste steuerungsrelevante offene Meldungen")
        show_table(stand, columns, max_rows=100)
        make_download(stand, "steuerungsrelevante_offene_langlaeufer.csv")

with tab6:
    st.subheader("Detaildaten")
    st.write("Hier sind alle gefilterten, relevanten Meldungen ohne MM und ohne Löschvermerke.")
    show_table(filtered_df, columns)
    make_download(filtered_df, "vat_kfz_servicemeldungen_gefiltert.csv")

# -----------------------------
# Diagnosebereich
# -----------------------------
with st.expander("Erkannte Spalten / Diagnose"):
    st.write("Diese Spalten hat die App automatisch erkannt:")
    st.json(columns)
    st.write("Alle Spalten im Excel:")
    st.write(list(raw_df.columns))
    st.write("Statusgruppen nach SAP-Systemstatus:")
    if columns.get("status"):
        status_diag = prepared_df.groupby([columns["status"], "_Statusgruppe"]).size().reset_index(name="Anzahl")
        st.dataframe(status_diag, use_container_width=True, hide_index=True)

    st.write("Wartungslogik Diagnose:")
    wartung_diag = {
        "Wartung_geplant_GewEnde_in_Zukunft": int(prepared_df["_Wartung_geplant"].sum()) if "_Wartung_geplant" in prepared_df.columns else 0,
        "Wartung_naechste_30_Tage": int(prepared_df["_Wartung_naechste_30_Tage"].sum()) if "_Wartung_naechste_30_Tage" in prepared_df.columns else 0,
        "Wartung_im_Verzug": int(prepared_df["_Wartung_im_Verzug"].sum()) if "_Wartung_im_Verzug" in prepared_df.columns else 0,
        "Wartung_ohne_Termin_offen": int(prepared_df["_Wartung_ohne_Termin_offen"].sum()) if "_Wartung_ohne_Termin_offen" in prepared_df.columns else 0,
    }
    st.write(wartung_diag)

    st.write("Ausgeschlossene Datensätze:")
    st.write({
        "MM_Meldungsarten": int(prepared_df["_Ausgeschlossen"].sum()) if "_Ausgeschlossen" in prepared_df.columns else 0,
        "Loeschvermerke": int(prepared_df["_Loeschvermerk"].sum()) if "_Loeschvermerk" in prepared_df.columns else 0,
    })
