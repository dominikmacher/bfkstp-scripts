````markdown
# 🏆 FDISK Bezirkswertung

PowerShell-Skript zur automatischen Erstellung einer **Bezirkswertung** aus FDISK‑CSV-Dateien.

## ✅ Features
- Einlesen mehrerer CSV-Dateien
- Kategorie-Auswertung (z. B. WertKlasse)
- Filter ungültiger Gruppen
- Ausschluss schlechter Ergebnisse
- Excel-Export (mit Rang & Markierungen)
- LOG-Datei mit Status aller Gruppen

## 📦 Voraussetzungen
- PowerShell 5+
- Modul: `ImportExcel`

```powershell
Install-Module ImportExcel -Scope CurrentUser
````

## ▶️ Nutzung

```powershell
.\auswertung_final.ps1
```

## 📁 Struktur

```
fdisk_export/   # CSV-Dateien
auswertung_final.ps1
```

## ⚙️ Wichtige Parameter

```powershell
$MinBewerbe = 3
$excludeWorstResults = 1
$CategoryCol = "WertKlasse"
```

## 📄 Output

* `Auswertung_je_Kategorie.xlsx`
* `Auswertung_LOG.txt`

***

🚒 Für Feuerwehrbewerbe optimiert

```
```
