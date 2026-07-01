<#
.SYNOPSIS
    Erstellt eine Bezirkswertung aus mehreren FDISK‑CSV‑Dateien (Ergebnislisten der Abschnittsbewerbe).

.DESCRIPTION
    Liest alle CSV‑Dateien aus dem Ordner "fdisk_export" ein, filtert ungültige Gruppen
    (WertGrp enthält "verschiedene"), berechnet die Ergebnisse je Bewerbsgruppe und Kategorie,
    erstellt eine Excel‑Auswertung und erzeugt eine LOG‑Datei.

    Funktionen:
    - CSV‑Einlesen und Gruppierung nach Kategorie
    - fehlende Ergebnisse werden als 0 gewertet
    - schlechteste Ergebnisse werden ausgeschlossen (konfigurierbar)
    - Excel‑Auswertung mit Rang, Formeln und Markierungen
    - LOG‑Datei mit alphabetischer Liste aller Gruppen und Status ("OK" / "zu wenig Bewerbsteilnahmen")

.PARAMETER MinBewerbe
    Mindestanzahl an Bewerben, damit eine Gruppe gewertet wird.

.PARAMETER filterTopGroups
    Anzahl der Top‑Gruppen pro Kategorie (0 = alle).

.PARAMETER excludeWorstResults
    Anzahl der auszuschließenden schlechtesten Ergebnisse.

.PARAMETER CategoryCol
    CSV‑Spalte, die die Kategorie definiert (z. B. "WertKlasse").

.INPUTS
    CSV‑Dateien im FDISK‑Format.

.OUTPUTS
    - Auswertung_je_Kategorie.xlsx
    - Auswertung_LOG.txt

.EXAMPLE
    PS C:\> .\auswertung_final.ps1
    Führt die komplette Bezirkswertung durch.

.NOTES
    Benötigt das PowerShell‑Modul "ImportExcel".
#>






# ------------------------------------------------------------
# KONFIGURATION
# ------------------------------------------------------------
$basePath            = Get-Location
$inputFolder         = Join-Path $basePath "fdisk_export"
$outputFile          = Join-Path $basePath "Auswertung_je_Kategorie.xlsx"

$MinBewerbe          = 3                  # Mindestanzahl an Bewerben pro Kategorie
$filterTopGroups     = 0                  # Anzahl der Top-Gruppen pro Kategorie (0 = alle)
$excludeWorstResults = 1                  # 0 = alle zählen, 1 = schlechtestes raus, 2 = zwei schlechteste raus
$CategoryCol         = "WertKlasse"       # Kategorie-Spalte im CSV

# ------------------------------------------------------------
# CSV-Dateien einlesen (mit korrektem Encoding!)
# ------------------------------------------------------------
$files = Get-ChildItem -Path $inputFolder -Filter *.csv
$categories = @{}

foreach ($file in $files) {
    $csv = Import-Csv $file.FullName -Delimiter ';' -Encoding Default

    foreach ($row in $csv) {

        # Bewerbsgruppen mit WertGrp "verschiedene" ignorieren
        if ($row.WertGrp -match "(?i)verschiedene") {
            continue
        }

        $gruppe    = $row.Gruppenname
        $kategorie = $row.$CategoryCol
        $gesamt    = $row.Gesamt

        if (-not $categories.ContainsKey($kategorie)) {
            $categories[$kategorie] = @{}
        }

        if (-not $categories[$kategorie].ContainsKey($gruppe)) {
            $categories[$kategorie][$gruppe] = [ordered]@{
                Gruppenname = $gruppe
                Bewerbe     = @{}
            }
        }

        $categories[$kategorie][$gruppe].Bewerbe[$file.Name] = $gesamt
    }
}


# ------------------------------------------------------------
# LOG-Datei erzeugen
# ------------------------------------------------------------
$logFile = Join-Path $basePath "Auswertung_LOG.txt"
if (Test-Path $logFile) { Remove-Item $logFile }

# Alle Gruppen sammeln
$allGroups = @()

foreach ($kategorie in $categories.Keys) {
    foreach ($gruppe in $categories[$kategorie].Keys) {

        $bewerbeCount = $categories[$kategorie][$gruppe].Bewerbe.Count
        $isRelevant   = ($bewerbeCount -ge $MinBewerbe)

        $entry = [PSCustomObject]@{
            Gruppenname = $gruppe
            Kategorie   = $kategorie
            Bewerbe     = $bewerbeCount
            Relevant    = $isRelevant
        }

        $allGroups += $entry
    }
}

# Alphabetisch sortieren
$allGroups = $allGroups | Sort-Object Gruppenname

# LOG schreiben
foreach ($g in $allGroups) {

    if ($g.Relevant) {
        $line = "{0} ({1}) - OK ({2} Bewerbe)" -f $g.Gruppenname, $g.Kategorie, $g.Bewerbe
    } else {
        $line = "{0} ({1}) - zu wenig Bewerbsteilnahmen ({2} Bewerbe)" -f $g.Gruppenname, $g.Kategorie, $g.Bewerbe
    }

    Add-Content -Path $logFile -Value $line
}

Write-Host "LOG geschrieben: $logFile"





# ------------------------------------------------------------
# Hilfsfunktion: Ergebnisobjekt erzeugen
# ------------------------------------------------------------
function New-GroupResultObject {
    param(
        [string]$Kategorie,
        [hashtable]$GroupData,
        [System.IO.FileInfo[]]$Files
    )

    # Reihenfolge: Gruppenname → Gesamt-Ergebnis-auto-berechnet → Gesamt-Ergebnis → Ergebnisse
    $obj = [ordered]@{
        Gruppenname               = $GroupData.Gruppenname
        "Gesamt-Ergebnis-auto-berechnet"         = 0
        "Gesamt-Ergebnis"  = ""   # Formel wird später gesetzt
    }

    $values = @()

    foreach ($file in $Files) {
        $colName = "Ergebnis $($file.Name)"

        if ($GroupData.Bewerbe.ContainsKey($file.Name)) {
            $raw = $GroupData.Bewerbe[$file.Name]
            if ($raw -ne $null -and $raw -ne "") {
                $value = [decimal]($raw -replace ',', '.')
            } else {
                $value = 0
            }
        } else {
            $value = 0
        }

        $obj[$colName] = $value
        $values += $value
    }

    # schlechteste 0/1/2 Ergebnisse ausschließen (für Code-Summe)
    if ($values.Count -gt 0) {
        $n = [Math]::Min($excludeWorstResults, $values.Count - 1)

        $sorted = $values | Sort-Object
        $use    = $sorted[$n..($sorted.Count - 1)]
        $sum    = ($use | Measure-Object -Sum).Sum
    } else {
        $sum = 0
    }

    $obj["Gesamt-Ergebnis-auto-berechnet"] = [decimal]$sum

    return [PSCustomObject]$obj
}

# ------------------------------------------------------------
# Excel-Datei vorbereiten
# ------------------------------------------------------------
if (Test-Path $outputFile) {
    Remove-Item $outputFile
}

# ------------------------------------------------------------
# Verarbeitung pro Kategorie
# ------------------------------------------------------------
foreach ($kategorie in $categories.Keys) {

    $groups = $categories[$kategorie]

    $eligible = $groups.Values | Where-Object {
        $_.Bewerbe.Count -ge $MinBewerbe
    }

    if ($eligible.Count -eq 0) { continue }

    $results = foreach ($g in $eligible) {
        New-GroupResultObject -Kategorie $kategorie -GroupData $g -Files $files
    }

    # Sortierung nach Gesamt-Ergebnis-auto-berechnet (Code)
    $results = $results | Sort-Object -Property "Gesamt-Ergebnis-auto-berechnet" -Descending

    if ($filterTopGroups -gt 0) {
        $results = $results | Select-Object -First $filterTopGroups
    }

    $sheetName = ($kategorie -replace '[^\w\s-]', '')
    $sheetName = $sheetName.Replace('mit', 'm.')
    $sheetName = $sheetName.Replace('ohne', 'o.')
    $sheetName = $sheetName.Substring(0, [Math]::Min(30, $sheetName.Length))

    $results | Export-Excel -Path $outputFile -WorksheetName $sheetName -AutoSize -Append
}

# ------------------------------------------------------------
# Excel-Nachbearbeitung
# ------------------------------------------------------------
if (Test-Path $outputFile) {
        
    $excel = Open-ExcelPackage -Path $outputFile

    foreach ($ws in $excel.Workbook.Worksheets) {

        # Headlines fett
        $ws.Cells[1,1,1,$ws.Dimension.End.Column].Style.Font.Bold = $true

        # Ergebnis-Spalten finden
        $ergebnisCols = @()
        for ($col = 1; $col -le $ws.Dimension.End.Column; $col++) {
            if ($ws.Cells[1, $col].Text -like "Ergebnis*") {
                $ergebnisCols += $col
            }
        }

        if ($ergebnisCols.Count -eq 0) { continue }

        # Spaltenpositionen
        $sumCol     = 2
        $formulaCol = 3

        # Format auf 2 Nachkommastellen
        foreach ($col in $ergebnisCols) {
            $ws.Cells[2, $col, $ws.Dimension.End.Row, $col].Style.Numberformat.Format = "0.00"
        }
        $ws.Cells[2, $sumCol,     $ws.Dimension.End.Row, $sumCol].Style.Numberformat.Format     = "0.00"
        $ws.Cells[2, $formulaCol, $ws.Dimension.End.Row, $formulaCol].Style.Numberformat.Format = "0.00"

        # schlechteste Ergebnisse gelb markieren + Excel-Formel setzen
        for ($row = 2; $row -le $ws.Dimension.End.Row; $row++) {

            $values = @()
            foreach ($col in $ergebnisCols) {
                $raw = $ws.Cells[$row, $col].Value

                if ($raw -eq $null -or $raw.ToString().Trim() -eq "") {
                    $num = 0
                } else {
                    $num = [decimal]($raw.ToString().Trim() -replace ',', '.')
                }

                $values += [PSCustomObject]@{
                    Col = $col
                    Val = $num
                }
            }

            if ($values.Count -eq 0) { continue }

            $n = [Math]::Min($excludeWorstResults, $values.Count - 1)
            $excluded = $values | Sort-Object Val | Select-Object -First $n

            foreach ($item in $excluded) {
                $ws.Cells[$row, $item.Col].Style.Fill.PatternType = "Solid"
                $ws.Cells[$row, $item.Col].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Yellow)
            }

            # Excel-Formel für Gesamt-Ergebnis-auto-berechnet-Formel setzen
            $includeCells = @()
            foreach ($v in $values) {
                if ($excluded.Col -notcontains $v.Col) {
                    $includeCells += $ws.Cells[$row, $v.Col].Address
                }
            }

            if ($includeCells.Count -gt 0) {
                $formula = "=SUM(" + ($includeCells -join ",") + ")"
                $ws.Cells[$row, $formulaCol].Formula = $formula
            }
        }

        # --------------------------------------------------------
        # Spalte "Gesamt-Ergebnis-auto-berechnet" entfernen
        # --------------------------------------------------------
        #$ws.DeleteColumn(2)


        # --------------------------------------------------------
        # Spalte "Gesamt-Ergebnis-auto-berechnet" in Rang-Spalte umwandeln
        # --------------------------------------------------------
        $ws.Cells[1,2].Value = "Rang"
        $ws.Cells[1,2].Style.Font.Bold = $true
        
        # Auto-Breite für Rang-Spalte
        $ws.Column(2).AutoFit()

        # Rang aufsteigend setzen
        $rank = 1
        for ($row = 2; $row -le $ws.Dimension.End.Row; $row++) {
            $ws.Cells[$row,2].Value = $rank
            $ws.Cells[$row, 2].Style.Numberformat.Format = "0"
            $rank++
        }
    }

    Close-ExcelPackage $excel

    Write-Host "Fertig! Excel wurde erstellt: $outputFile"
}
else {
    Write-Host "Keine Excel-Datei geschrieben!!"
}

Write-Host ""
