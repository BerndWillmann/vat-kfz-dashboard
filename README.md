# VAT KFZ Werkstätten-Cockpit

Streamlit-Dashboard zur Auswertung von SAP-Servicemeldungen aus Excel.

## Funktionen

- MM-Meldungen werden ausgeschlossen
- MF-Reparaturen werden als wichtigste Meldungen separat ausgewertet
- MF-Reparaturen länger als 10 Tage offen werden hervorgehoben
- WE/WK/W1/W2/W3/WZ/WP werden als Wartungen ausgewertet
- Wartungen im Verzug werden angezeigt
- G/GM werden als Garantie ausgewertet
- F-Meldungen werden separat dargestellt
- Auswertung nach Technischem Platz / Equipment
- Top-Störer und Langläufer
- Detailtabelle mit CSV-Export

## Dateien

- `app.py` - Streamlit-App
- `requirements.txt` - benötigte Python-Pakete
- optional: `Servicemeldungen_VAT_KFZ.xlsx` - Excel-Datei direkt im Repository ablegen

## Start lokal

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Nutzung

Entweder:

1. Excel-Datei in der App hochladen

oder:

2. Excel-Datei mit dem Namen `Servicemeldungen_VAT_KFZ.xlsx` direkt in das GitHub-Repository legen.

## Wichtige Logik

- Ausschluss: Meldungsart beginnt mit `MM`
- Reparatur: Meldungsart beginnt mit `MF`
- Wartung: Meldungsart beginnt mit `WE`, `WK`, `W1`, `W2`, `W3`, `WZ`, `WP`
- Garantie: Meldungsart beginnt mit `G` oder `GM`
- F-Meldung: Meldungsart beginnt mit `F`, aber nicht `MF`

## Hinweise

Falls deine SAP-Excel-Spalten andere Namen haben, erkennt die App viele Varianten automatisch. Im Bereich "Erkannte Spalten / Diagnose" siehst du, welche Spalten gefunden wurden.
