# LibreOffice Calc Export — Design Spec

**Datum:** 2026-04-10
**Status:** Genehmigt

## Zusammenfassung

Die bestehende `AuswertungForm` (Pivot-Dialog, Filter, mehrere Export-Formate, DB-Pfadliste) wird komplett entfernt und durch einen einzelnen Toolbar-Button ersetzt: **"In LibreOffice Calc öffnen"**.

Ein Klick exportiert die 3 Datenbank-VIEWs als XLSX-Datei mit je einem Sheet und öffnet sie in LibreOffice Calc.

## Anforderungen

- Ein einziger Button in der Toolbar (neben "Aktualisieren")
- Export der 3 VIEWs: `Inventar_Geraete`, `Inventar_Software`, `Inventar_Ports`
- Jede VIEW = ein Sheet in der XLSX-Datei
- Fortschrittsanzeige im Button-Text während des Exports
- Automatisches Öffnen in LibreOffice Calc
- Fehlermeldung falls LibreOffice nicht gefunden wird (mit Dateipfad)
- AuswertungForm.cs wird gelöscht

## Architektur

### Toolbar-Button

- Label: `📊 In LibreOffice Calc öffnen`
- Position: rechts neben dem bestehenden `🔄 Aktualisieren`-Button
- Styling: konsistent mit bestehendem Aktualisieren-Button

### Export-Ablauf

1. Button disabled, Text → "Exportiere... (1/3)"
2. `SELECT * FROM Inventar_Geraete` → Sheet "Geräte"
3. Text → "Exportiere... (2/3)"
4. `SELECT * FROM Inventar_Software` → Sheet "Software"
5. Text → "Exportiere... (3/3)"
6. `SELECT * FROM Inventar_Ports` → Sheet "Ports"
7. XLSX speichern: `%TEMP%\Inventarisierung_<yyyy-MM-dd_HHmmss>.xlsx`
8. LibreOffice Calc suchen → `scalc.exe` starten mit XLSX-Pfad
9. Button re-enabled, Text zurücksetzen

### LibreOffice-Suche

5 bekannte Installationspfade (wie bisher im Code) + PATH-Fallback:
- `C:\Program Files\LibreOffice\program\scalc.exe`
- `C:\Program Files (x86)\LibreOffice\program\scalc.exe`
- `C:\Program Files\LibreOffice 7\program\scalc.exe`
- `C:\Program Files\LibreOffice 24\program\scalc.exe`
- `C:\Program Files\LibreOffice 25\program\scalc.exe`
- Fallback: `scalc` via PATH

### Fehlerbehandlung

| Fall | Reaktion |
|------|----------|
| Keine Daten in DB | MessageBox: "Keine Daten vorhanden." |
| LibreOffice nicht gefunden | MessageBox: "LibreOffice nicht gefunden. Die Datei wurde gespeichert unter: {Pfad}" |
| DB-Fehler | MessageBox mit Fehlermeldung |

## Betroffene Dateien

| Datei | Aktion |
|-------|--------|
| `Program.cs` | Neuer Button + `ExportAndOpenInCalc()` Methode |
| `DatabaseManager.cs` | Neue Methode `GetViewData(string viewName)` → `DataTable` |
| `AuswertungForm.cs` | Wird gelöscht |
| `Inventarisierung.csproj` | Referenz auf AuswertungForm entfernen |

## Abhängigkeiten

- **ClosedXML** (bereits vorhanden) — XLSX-Erzeugung
- Keine neuen NuGet-Pakete erforderlich

## Nicht im Scope

- PDF-Export (entfällt mit AuswertungForm)
- CSV-Export (entfällt mit AuswertungForm)
- Pivot-Gruppierung / Filter (entfällt)
- LibreOffice Base Integration (verworfen — braucht JDBC-Treiber)
