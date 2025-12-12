param(
    # Site URL
    [string]$SiteUrl = "https://intranet.volvocars.net/sites/TestZone",

    # Naam van de documentbibliotheek
    [string]$LibraryName = "Documents",

    # BRONMAP: enkel hier (en subfolders) uit moven
    [string]$SourceFolderRelativeUrl = "/sites/TestZone/Shared Documents/BOLD",

    # DOELMAP: alles komt hier vlak te staan
    [string]$TargetFolderRelativeUrl = "/sites/TestZone/Shared Documents/FLATTENED",

    # Testmodus: toon wat er zou gebeuren, maar voer geen Move uit
    [switch]$WhatIf
)


# ------------------------------------------------------------
# 1) Doelmap controleren / aanmaken
# ------------------------------------------------------------
Write-Host "Controleren of doelmap bestaat: $TargetFolderRelativeUrl ..."
$targetFolder = Get-PnPFolder -Url $TargetFolderRelativeUrl -ErrorAction SilentlyContinue

if (-not $targetFolder) {
    Write-Host "Doelmap bestaat niet, aanmaken..."

    $parentUrl  = Split-Path $TargetFolderRelativeUrl -Parent
    $folderName = Split-Path $TargetFolderRelativeUrl -Leaf

    if (-not $WhatIf) {
        New-PnPFolder -Name $folderName -Folder $parentUrl | Out-Null
    }

    $targetFolder = Get-PnPFolder -Url $TargetFolderRelativeUrl
}

Write-Host "Doelmap OK: $TargetFolderRelativeUrl"
Write-Host "Bronmap:   $SourceFolderRelativeUrl"

# ------------------------------------------------------------
# 2) Enkel items ophalen onder de BRONMAP (sneller + correct)
# ------------------------------------------------------------
Write-Host "Ophalen documenten onder bronmap '$SourceFolderRelativeUrl' ..."
$list = Get-PnPList -Identity $LibraryName

$items = Get-PnPListItem -List $list -PageSize 500 `
    -FolderServerRelativeUrl $SourceFolderRelativeUrl `
    -Fields "FileLeafRef","FileRef","FSObjType" `
    -ScriptBlock { param($batch) $batch }

Write-Host "Aantal items gevonden onder bronmap: $($items.Count)"

# Enkel bestanden, geen folders, en niet al in de doelmap
$filesToMove = $items | Where-Object {
    $_["FSObjType"] -eq 0 -and
    $_["FileRef"] -notlike "$TargetFolderRelativeUrl/*"
}

Write-Host "Aantal bestanden die verplaatst zullen worden: $($filesToMove.Count)"

# ------------------------------------------------------------
# 3) Flatten + duplicates veilig behandelen
# ------------------------------------------------------------
$counter = 0

foreach ($item in $filesToMove) {
    $counter++

    $fileName  = $item["FileLeafRef"]
    $sourceUrl = $item["FileRef"]

    # Extra beveiliging: nooit buiten bronmap verplaatsen
    if ($sourceUrl -notlike "$SourceFolderRelativeUrl/*") {
        Write-Host "SKIP (buiten bronmap): $sourceUrl" -ForegroundColor Yellow
        continue
    }

    $baseName  = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    $extension = [System.IO.Path]::GetExtension($fileName)

    Write-Host ""
    Write-Host "[$counter/$($filesToMove.Count)] Verwerken: $sourceUrl"

    # Standaard doel-URL
    $targetFileUrl = "$TargetFolderRelativeUrl/$fileName"

    # Check op duplicates in de doelmap -> hernoem naar _DUP1, _DUP2, ...
    $dupIndex = 1
    $finalFileName = $fileName

    while (Get-PnPFile -Url $targetFileUrl -ErrorAction SilentlyContinue) {
        $finalFileName = "{0}_DUP{1}{2}" -f $baseName, $dupIndex, $extension
        $targetFileUrl = "$TargetFolderRelativeUrl/$finalFileName"
        $dupIndex++
    }

    if ($WhatIf) {
        if ($finalFileName -ne $fileName) {
            Write-Host "  WOULD MOVE (DUPLICATE-NAAM):" -ForegroundColor Yellow
        } else {
            Write-Host "  WOULD MOVE:" -ForegroundColor Yellow
        }
        Write-Host "  From: $sourceUrl"
        Write-Host "  To:   $targetFileUrl"
        continue
    }

    # Verplaatsen
    try {
        Move-PnPFile -ServerRelativeUrl $sourceUrl -TargetUrl $targetFileUrl -Force -AllowSchemaMismatch -IgnoreVersionHistory:$false -ErrorAction Stop | Out-Null

        if ($finalFileName -ne $fileName) {
            Write-Host "  VERPLAATST (DUPLICATE-NAAM):"
        } else {
            Write-Host "  VERPLAATST:"
        }
        Write-Host "  From: $sourceUrl"
        Write-Host "  To:   $targetFileUrl"
    }
    catch {
        Write-Host "  FOUT bij verplaatsen:" -ForegroundColor Red
        Write-Host "  From: $sourceUrl"
        Write-Host "  To:   $targetFileUrl"
        Write-Host $_
        continue
    }
}

Write-Host ""
Write-Host "Klaar. Verwerkte bestanden: $counter"
