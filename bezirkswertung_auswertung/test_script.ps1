Install-Module -Name ImportExcel -Scope CurrentUser



# Encoding fix (wichtig für Umlaute!)
#[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Import-Module ImportExcel

$inputFolder = "fdisk_export"
$outputFile  = "Top3_je_Kategorie.xlsx"

# Dateien laden
$files = Get-ChildItem -Path $inputFolder -Filter *.xlsx

$allData = @()

foreach ($file in $files) {
    Write-Host "Lese: $($file.Name)"
    
    $data = Import-Excel -Path $file.FullName

    if ($data) {
        $allData += $data
    }
    Write-Host ($data | Format-Table | Out-String)

}

# Datentypen sicherstellen
$allData | ForEach-Object {
    #$_.Gesamt = [double]($_.Gesamt -replace ",",".")
}

# Daten vorbereiten (Umlaute sicher)
$cleanData = foreach ($row in $allData) {

    # Werte sauber auslesen (Umlaute fix!)
    $punkte = $row.PSObject.Properties["Punkte-Löschangriff"].Value
    $fehler = $row.PSObject.Properties["Fehler-Löschangriff"].Value
    $klasse = $row.PSObject.Properties["WertKlasseKurz"].Value

Write-Host "_____: " + $row.PSObject.Properties["Punkte-Löschangriff"].Value
#Write-Host ($row | Format-Table | Out-String)

    # Kategorien zusammenlegen
    if ($klasse -in @("BA","BAG1","BAO","BAW")) {
        $klasse = "BA"
    }
    elseif ($klasse -in @("BBG1","BBW")) {
        $klasse = "BB"
    }
    elseif ($klasse -in @("SA","SAG1","SAO","SAW")) {
        $klasse = "SA"
    }
    elseif ($klasse -in @("SBG1")) {
        $klasse = "SB"
    }

    [PSCustomObject]@{
        Kategorie                = $klasse
        Gruppenname              = $row.Gruppenname
        Gesamt                   = [double]$row.Gesamt
        Punkte_Loeschangriff     = [double]$punkte
        Fehler_Loeschangriff     = [double]$fehler
        Bundesland               = $row.Bundesland
    }
}

# Gruppieren
$groups = $cleanData | Group-Object Kategorie

# Ergebnis-Datei neu erzeugen
if (Test-Path $outputFile) {
    Remove-Item $outputFile
}

foreach ($group in $groups) {

    $sheetName = $group.Name

    if ($group.Name -in @("BAG2","SAG2","BBG2","SBG2")) {
        Write-Host "Ignoriere: $sheetName"
        continue   
    }

    # Excel erlaubt max. 31 Zeichen
    if ($sheetName.Length -gt 31) {
        $sheetName = $sheetName.Substring(0,31)
    }

    Write-Host "Erstelle Tabellenblatt: $sheetName"

    $top3 = $group.Group |
        Sort-Object -Property @{Expression="Gesamt";Descending=$true} |
        Select-Object -First 3

    $platz = 1

    $result = foreach ($team in $top3) {
        [PSCustomObject]@{
            Platz               = $place
            Gruppenname         = $team.Gruppenname
            Gesamt              = $team.Gesamt
            Punkte_Loeschangriff= $team.Punkte_Loeschangriff
            Fehler_Loeschangriff= $team.Fehler_Loeschangriff
            Bundesland          = $team.Bundesland
        }
        $platz++
    }

    # In eigenes Tabellenblatt schreiben
    $result | Export-Excel `
        -Path $outputFile `
        -WorksheetName $sheetName `
        -AutoSize `
        -BoldTopRow `
        -Append
}

Write-Host "✅ Fertig! Datei: $outputFile"