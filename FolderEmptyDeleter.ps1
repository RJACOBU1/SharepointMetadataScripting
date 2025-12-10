param(
    # Naam van de documentbibliotheek
    [string]$listName = "Documents",

    # Server-relative pad van de root folder
    [string]$rootFolderFilter = "/sites/TestZone/Shared Documents/010 Crossconveyor"

)

Write-Host "Start opruimen van lege mappen onder '$rootFolderFilter' in bibliotheek '$listName'" -ForegroundColor Cyan

$iteration     = 0
$totalDeleted  = 0

while ($true) {

    $iteration++
    Write-Host "------------------------------" -ForegroundColor DarkCyan
    Write-Host "Iteratie $iteration" -ForegroundColor DarkCyan

    # 1. Haal ALLE folders onder de root op, in één keer
    $allFolders = Get-PnPListItem -List $listName -PageSize 5000 `
        -Fields "FileRef","FileLeafRef","FSObjType","ItemChildCount","FolderChildCount" |
        Where-Object {
            $_["FSObjType"] -eq 1 -and                            # 1 = folder
            $_["FileLeafRef"] -ne "Forms" -and                    # Forms-folder negeren
            $_["FileRef"] -like "$rootFolderFilter*" -and         # onder jouw root
            $_["FileRef"] -ne $rootFolderFilter                   # root zelf nooit verwijderen
        }

    if (-not $allFolders) {
        Write-Host "Geen folders gevonden onder deze hoofdmap." -ForegroundColor Yellow
        break
    }

    Write-Host ("Totaal aantal mappen in scope: " + $allFolders.Count) -ForegroundColor Cyan

    # 2. Filter op echt lege mappen op basis van de child counts
    $emptyFolders = $allFolders | Where-Object {
        # defensief casten naar int, null => 0
        ([int]$_.FieldValues["ItemChildCount"])   -eq 0 -and
        ([int]$_.FieldValues["FolderChildCount"]) -eq 0
    }

    $count = $emptyFolders.Count

    if ($count -eq 0) {
        Write-Host "Geen lege mappen meer gevonden. Script klaar." -ForegroundColor Green
        break
    }

    Write-Host "$count lege map(pen) gevonden in iteratie $iteration." -ForegroundColor Yellow

    foreach ($item in $emptyFolders) {
        $folderUrl = $item.FieldValues["FileRef"]

        Write-Host "Verwijder lege map: $folderUrl" -ForegroundColor Red
        Remove-PnPListItem -List $listName -Identity $item.Id -Force
        $totalDeleted++

    }
}

Write-Host "Script afgelopen. Totaal aantal verwijderde mappen: $totalDeleted" -ForegroundColor Cyan
