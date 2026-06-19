# Voraussetzung: Modul ImportExcel  
#Install-Module -Name ImportExcel -Scope CurrentUser


Import-Module ImportExcel



# Aktueller Pfad (funktioniert im Skript und interaktiv)
$currentPath = Get-Location
# Ausgabe-Excel im aktuellen Pfad
$outputFile = Join-Path $currentPath "Top3_je_Kategorie.xlsx"

$csvPath = Join-Path $currentPath "fdisk_export"

# CSV-Dateien laden
$csvFiles = Get-ChildItem -Path $csvPath -Filter *.csv

# Gesamtdaten einlesen
$data = foreach ($file in $csvFiles) {
    #Import-Csv -Path $file.FullName -Delimiter ";" 
    #Import-Csv -Path $file.FullName -Delimiter ";" -Encoding Default

    $rows = Import-Csv -Path $file.FullName -Delimiter ";" -Encoding Default
    foreach ($row in $rows) {
        $row | Add-Member -NotePropertyName Quelldatei -NotePropertyValue $file.Name
        $row
    }
}

# Kategorien aus WertKlasse ermitteln
$categories = $data | Select-Object -ExpandProperty WertKlasse -Unique

# Falls Excel existiert → löschen
if (Test-Path $outputFile) { Remove-Item $outputFile }

foreach ($cat in $categories) {

    # Alle Einträge der Kategorie holen
    $catRows = $data | Where-Object { $_.WertKlasse -eq $cat }

    # PRO FEUERWEHR nur das beste Ergebnis behalten
    $bestPerFF = $catRows |
        Group-Object -Property Gruppenname |
        ForEach-Object {
            $_.Group | Sort-Object { [double]$_.Gesamt } -Descending | Select-Object -First 1
        }

    # Jetzt die Top 3 der Kategorie bestimmen
    $top3 = $bestPerFF |
        Sort-Object { [double]$_.Gesamt } -Descending |
        Select-Object -First 3

    # Rang Bezirkswertung hinzufügen
    $rank = 1
    foreach ($row in $top3) {
        $row | Add-Member -NotePropertyName "Rang Bezirkswertung" -NotePropertyValue $rank -Force
        $rank++
    }

    # Tabellenblattname bereinigen
    $sheetName = ($cat -replace '[^\w\s-]', '').Substring(0, [Math]::Min(28, $cat.Length))

    # Finale Spaltenreihenfolge
    $ordered = $top3 |
        Select-Object `
            "Rang Bezirkswertung",
            Gruppenname,
            Gesamt,
            @{ Name = 'Punkte-Loeschangriff'; Expression = {
                ($_.PSObject.Properties |
                    Where-Object { $_.Name -like 'Punkte*schangriff' }
                ).Value
            }},
            @{ Name = 'Fehler-Loeschangriff'; Expression = {
                ($_.PSObject.Properties |
                    Where-Object { $_.Name -like 'Fehler*schangriff' }
                ).Value
            }},
            Alterspunkte,
            WertGrp,
            WertKlasse,
            WertKlasseKurz,
            BewNr,
            Instanz,
            InstanzNr,
            Instanzart,
            AFKDO,
            Quelle

    # Export ins Excel
    $ordered |
        Export-Excel -Path $outputFile -WorksheetName $sheetName -AutoSize -Append
}


Write-Host "Fertig! Excel wurde erstellt: $outputFile"