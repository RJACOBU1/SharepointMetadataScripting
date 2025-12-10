param(
    # Documentbibliotheek
    [string]$LibraryName = "Documents",

    # Bronkolom (interne naam)
    [string]$SourceField = "EquipmentDocument",

    # Doelkolom (interne naam)
    [string]$TargetField = "CopyPlace",

    # Enkel updaten als de doelkolom leeg is?
    [switch]$OnlyWhenTargetEmpty,

    # Batchgrootte
    [int]$BatchSize = 200,

    # Testmodus
    [switch]$WhatIf
)

Import-Module PnP.PowerShell -ErrorAction Stop

Write-Host "Ophalen items uit '$LibraryName'..." -ForegroundColor Cyan
$items = Get-PnPListItem -List $LibraryName -PageSize 2000
Write-Host "Totaal aantal items: $($items.Count)" -ForegroundColor Green

$batch   = $null
$inBatch = 0
$updates = 0

if (-not $WhatIf) {
    $batch = New-PnPBatch
}

foreach ($item in $items) {

    $src = $item[$SourceField]
    $tgt = $item[$TargetField]

    # Geen bronwaarde → niets te kopiëren
    if (-not $src) { continue }

    # Optioneel: doel alleen vullen als die leeg is
    if ($OnlyWhenTargetEmpty -and $tgt) { continue }

    # Vergelijk bron en doel als tekst (werkt voor choice en text)
    $srcText = $src.ToString()
    $tgtText = if ($tgt) { $tgt.ToString() } else { "" }

    # Als al hetzelfde → overslaan
    if ($srcText -eq $tgtText) { continue }

    Write-Host "Item $($item.Id): '$SourceField'='$srcText' → '$TargetField' (was: '$tgtText')" -ForegroundColor Yellow
    $updates++

    if (-not $WhatIf) {
        Set-PnPListItem -List $LibraryName -Identity $item.Id -Values @{
            $TargetField = $src
        } -Batch $batch | Out-Null

        $inBatch++
        if ($inBatch -ge $BatchSize) {
            Write-Host "Batch van $BatchSize items versturen..." -ForegroundColor Cyan
            Invoke-PnPBatch -Batch $batch
            $batch = New-PnPBatch
            $inBatch = 0
        }
    }
}

if (-not $WhatIf -and $inBatch -gt 0) {
    Write-Host "Laatste batch van $inBatch items versturen..." -ForegroundColor Cyan
    Invoke-PnPBatch -Batch $batch
}

Write-Host "Aantal geplande updates: $updates" -ForegroundColor Green
Write-Host "Klaar." -ForegroundColor Green
