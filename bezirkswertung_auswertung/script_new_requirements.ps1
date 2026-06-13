# ------------------------------------------------------------
# KONFIGURATION
# ------------------------------------------------------------
$basePath            = Get-Location
$inputFolder         = Join-Path $basePath "fdisk_export"
$outputFile          = Join-Path $basePath "Top3_je_Kategorie.xlsx"

$MinBewerbe          = 4                 # Mindestanzahl an Bewerben pro Kategorie
$filterTopGroups     = 0                 # Anzahl der Top-Gruppen pro Kategorie (0 = alle)
$excludeWorstResults = 1                 # 0 = alle zählen, 1 = schlechtestes raus, 2 = zwei schlechteste raus
$CategoryCol         = "WertKlasse"      # Kategorie-Spalte im CSV

# ------------------------------------------------------------
# CSV-Dateien einlesen
# ------------------------------------------------------------
$files = Get-ChildItem -Path $inputFolder -Filter *.csv

# Struktur: Kategorie → Gruppe → Bewerbsergebnisse
$categories = @{}

foreach ($file in $files) {
    $csv = Import-Csv $file.FullName -Delimiter ';'

    foreach ($row in $csv) {
        $gruppe    = $row.Gruppenname
        $kategorie = $row.$CategoryCol
        $gesamt    = $row.Gesamt   # als String, wir rechnen später sauber

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
# Hilfsfunktion: Ergebnisobjekt für eine Gruppe erzeugen
# ------------------------------------------------------------
function New-GroupResultObject {
    param(
        [string]$Kategorie,
        [hashtable]$GroupData,
        [System.IO.FileInfo[]]$Files
    )

    # Reihenfolge: Gruppenname → Gesamt-Ergebnis → Ergebnisse
    $obj = [ordered]@{
        Gruppenname      = $GroupData.Gruppenname
        "Gesamt-Ergebnis"  = 0
    }

    $values = @()

    foreach ($file in $Files) {
        $colName = "Ergebnis $($file.Name)"
        $value   = $null

        if ($GroupData.Bewerbe.ContainsKey($file.Name)) {
            # Dezimal-Komma in Punkt umwandeln, dann in Zahl
            $raw = $GroupData.Bewerbe[$file.Name]
            if ($raw -ne $null -and $raw -ne "") {
                $value = [decimal]($raw -replace ',', '.')
                $values += $value
            }
        }

        $obj[$colName] = $value
    }

    # schlechteste 0/1/2 Ergebnisse ausschließen
    if ($values.Count -gt 0) {
        $n = [Math]::Min($excludeWorstResults, $values.Count - 1)
        if ($n -gt 0) {
            $sorted = $values | Sort-Object
            $use    = $sorted[$n..($sorted.Count - 1)]
            $sum    = ($use | Measure-Object -Sum).Sum
        } else {
            $sum = ($values | Measure-Object -Sum).Sum
        }
    } else {
        $sum = 0
    }

    $obj["Gesamt-Ergebnis"] = [decimal]$sum

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

    # Gruppen filtern: mindestens X Bewerbe in dieser Kategorie
    $eligible = $groups.Values | Where-Object {
        $_.Bewerbe.Count -ge $MinBewerbe
    }

    if ($eligible.Count -eq 0) {
        continue
    }

    # Ergebnisobjekte erzeugen
    $results = foreach ($g in $eligible) {
        New-GroupResultObject -Kategorie $kategorie -GroupData $g -Files $files
    }

    # Sortierung nach Gesamt-Ergebnis (absteigend)
    $results = $results | Sort-Object -Property "Gesamt-Ergebnis" -Descending

    # Top-Filter anwenden (0 = alle)
    if ($filterTopGroups -gt 0) {
        $results = $results | Select-Object -First $filterTopGroups
    }

    # Tabellenblattname bereinigen
    $sheetName = ($kategorie -replace '[^\w\s-]', '').Substring(0, [Math]::Min(28, $kategorie.Length))

    # In Excel schreiben
    $results | Export-Excel -Path $outputFile -WorksheetName $sheetName -AutoSize -Append
}

# ------------------------------------------------------------
# Schlechteste Ergebnisse in Excel gelb markieren
# ------------------------------------------------------------
$excel = Open-ExcelPackage -Path $outputFile

foreach ($ws in $excel.Workbook.Worksheets) {

    # Ergebnis-Spalten finden
    $ergebnisCols = @()
    for ($col = 1; $col -le $ws.Dimension.End.Column; $col++) {
        $title = $ws.Cells[1, $col].Text
        if ($title -like "Ergebnis*") {
            $ergebnisCols += $col
        }
    }

    if ($ergebnisCols.Count -eq 0) { continue }

    # Jede Zeile bearbeiten
    for ($row = 2; $row -le $ws.Dimension.End.Row; $row++) {

        $values = @()
        foreach ($col in $ergebnisCols) {
            $cell = $ws.Cells[$row, $col]
            if ($cell.Value -ne $null -and $cell.Value -ne "") {
                $values += [PSCustomObject]@{
                    Col = $col
                    Val = [decimal]($cell.Value.ToString() -replace ',', '.')
                }
            }
        }

        if ($values.Count -eq 0) { continue }

        $n = [Math]::Min($excludeWorstResults, $values.Count - 1)
        if ($n -le 0) { continue }

        $excluded = $values | Sort-Object Val | Select-Object -First $n

        foreach ($item in $excluded) {
            $ws.Cells[$row, $item.Col].Style.Fill.PatternType = "Solid"
            $ws.Cells[$row, $item.Col].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Yellow)
        }
    }
}

Close-ExcelPackage $excel

Write-Host "Fertig! Excel wurde erstellt: $outputFile"
Write-Host ""
