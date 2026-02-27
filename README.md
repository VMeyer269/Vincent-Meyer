# UiPath Excel Filter Bot (Invoke Code, C#)

Vergleich von zwei Excel-Monatsstaenden in UiPath mit `Invoke Code` (C#) auf Basis von `DataTable`.

## Inhalt

- [uipath/InvokeCode_ExcelAbgleich.cs](/tmp/Ui-Path-Excel-Filter-Bot/uipath/InvokeCode_ExcelAbgleich.cs)
- [uipath/README_UiPath_ExcelAbgleich.md](/tmp/Ui-Path-Excel-Filter-Bot/uipath/README_UiPath_ExcelAbgleich.md)

## Was der Bot erkennt

- `NEU`: Bilanzkreis nur in neuem Monat vorhanden
- `WEG`: Bilanzkreis nur im alten Monat vorhanden
- `RLM_NEU` / `RLM_WEG`: Typwechsel innerhalb bestehender Bilanzkreise
- `SLP_NEU` / `SLP_WEG`: Typwechsel innerhalb bestehender Bilanzkreise
- `WECHSEL_TAGES_STUNDEN`: Regimewechsel zwischen Tages- und Stundenlogik

## Technische Highlights

- Robustes Spaltenhandling ohne `ItemArray`-Fehler
- Case-insensitive Erkennung (`RLM`/`SLP`)
- Optionales Deduplizieren und Sortieren der Ergebnisdaten
- Direkt nutzbar in UiPath `Invoke Code` (In/Out `DataTable`)

## Quick Start in UiPath

1. `Invoke Code` Activity einfuegen (`CSharp`)
2. In-Arguments: `dtAlt`, `dtNeu` (`System.Data.DataTable`)
3. Out-Argument: `dtChanges` (`System.Data.DataTable`)
4. Code aus `uipath/InvokeCode_ExcelAbgleich.cs` einfuegen
5. Namespaces: `System`, `System.Data`, `System.Linq`, `System.Collections.Generic`

## Hinweis

Bitte vor produktivem Einsatz interne Begriffe und ggf. sensible Daten anonymisieren.
