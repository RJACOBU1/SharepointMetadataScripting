param(
    # Documentbibliotheek
    [string]$LibraryName = "Documents",

    # Bron-lookup kolom (interne naam)
    [string]$SourceField = "StationNumbers",

    # Doel-lookup kolom (interne naam)
    [string]$TargetField = "StationNumberEff",

    # Titel van de lijst waar de DOEL-lookup naar verwijst
    [string]$TargetLookupListTitle = "LOCATIONS_Maximo",

    # Veld in de doellijst waarop we matchen (meestal Title)
    [string]$TargetLookupKeyField = "Title",

    # Testmodus
    [switch]$WhatIf
)

Write-Host "Ophalen items uit doellijst '$TargetLookupListTitle'..." -ForegroundColor Cyan
$targetLookupItems = Get-PnPListItem -List $TargetLookupListTitle -PageSize 2000

# Maak een mapping: sleutel = tekst (Title), waarde = ID in doellijst
$lookupMap = @{}
foreach ($li in $targetLookupItems) {
    $key = $li[$TargetLookupKeyField]
    if ($key) {
        $lookupMap[$key.ToString()] = $li.Id
    }
}

Write-Host "Aantal items in doellijst: $($lookupMap.Count)" -ForegroundColor Green

Write-Host "Ophalen documenten uit bibliotheek '$LibraryName'..." -ForegroundColor Cyan
$items = Get-PnPListItem -List $LibraryName -PageSize 2000

Write-Host "Aantal gevonden items in bibliotheek: $($items.Count)" -ForegroundColor Green

foreach ($item in $items) {

    $src = $item[$SourceField]

    # Geen bronwaarde → niets doen
    if (-not $src) { continue }

    # Bron is een single lookup (FieldLookupValue)
    $srcLookupValue = $src.LookupValue

    if (-not $srcLookupValue) { continue }

    # Kijk of dezelfde tekst bestaat in de doellijst
    if ($lookupMap.ContainsKey($srcLookupValue)) {
        $targetId = $lookupMap[$srcLookupValue]

        Write-Host "Item $($item.Id): '$SourceField'='$srcLookupValue' → '$TargetField' (ID $targetId)" -ForegroundColor Yellow

        if (-not $WhatIf) {
            Set-PnPListItem -List $LibraryName -Identity $item.Id -Values @{
                $TargetField = $targetId
            } | Out-Null
        }

    } else {
        Write-Warning "Geen match gevonden in doellijst voor '$srcLookupValue' (item $($item.Id))"
    }
}
