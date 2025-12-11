param(
    # Site URL
    [string]$SiteUrl = "https://intranet.volvocars.net/sites/TestZone",

    # Naam van de documentbibliotheek
    [string]$LibraryName = "Documents",

    [string]$SourceFolderRelativeUrl = "/sites/TestZone/Shared Documents/010 Crossconveyor",

    # Doelmap (server-relative URL) waar alles vlak komt te staan
    # Bijv.: "/sites/TestZone/Shared Documents/FLATTENED"
    [string]$TargetFolderRelativeUrl = "/sites/TestZone/Shared Documents/FLATTENED",

    # Testmodus: toon wat er zou gebeuren, maar voer geen Move uit
    [switch]$WhatIf
)


### 2. Controleren / aanmaken doelmap
Write-Host "Controleren of doelmap bestaat: $TargetFolderRelativeUrl ..."
$targetFolder = Get-PnPFolder -Url $TargetFolderRelativeUrl -ErrorAction SilentlyContinue

if (-not $targetFolder) {
    Write-Host "Doelmap bestaat niet, aanmaken..."

    $parentUrl = Split-Path $TargetFolderRelativeUrl -Parent
    $folderName = Split-Path $TargetFolderRelativeUrl -Leaf

    if (-not $WhatIf) {
        New-PnPFolder -Name $folderName -Folder $parentUrl | Out-Null
    }

    $targetFolder = Get-PnPFolder -Url $TargetFolderRelativeUrl
}

Write-Host "Doelmap OK: $TargetFolderRelativeUrl"

### 3. Alle items ophalen
Write-Host "Ophalen van alle items in lijst '$LibraryName' ..."
$list = Get-PnPList -Identity $LibraryName

$items = Get-PnPListItem -List $list -PageSize 500 `
    -Fields "FileLeafRef","FileRef","FSObjType" `
    -ScriptBlock { param($batch) $batch }

Write-Host "Totaal aantal items in lijst: $($items.Count)"

# Enkel bestanden, geen folders, en niet al in de doelmap
$filesToMove = $items | Where-Object {
    $_["FSObjType"] -eq 0 -and
    $_["FileRef"] -notlike "$TargetFolderRelativeUrl/*"
}

Write-Host "Aantal bestanden die verplaatst zullen worden: $($filesToMove.Count)"

### 4. Flatten + duplicates veilig behandelen

function Set-DuplicateMetaSafely {
    param(
        [Microsoft.SharePoint.Client.ListItem]$Item,
        [string]$DuplicateGroupKey,
        [string]$OriginalPath
    )

    # Deze functie probeert metadata te zetten, maar faalt stil als kolommen niet bestaan
    try {
        $values = @{}
        if ($DuplicateGroupKey) { $values["DuplicateGroupKey"] = $DuplicateGroupKey }
        if ($OriginalPath)      { $values["OriginalPath"]      = $OriginalPath }

        if ($values.Count -gt 0) {
            Set-PnPListItem -List $LibraryName -Identity $Item.Id -Values $values -ErrorAction Stop | Out-Null
        }
    }
    catch {
        # Kolom bestaat niet of andere meta-fout: we negeren dit om het script niet te breken
        write-Host "geen path en groupkey"
    }
}

$counter = 0

foreach ($item in $filesToMove) {
    $counter++

    $fileName = $item["FileLeafRef"]
    $sourceUrl = $item["FileRef"]

    $baseName  = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    $extension = [System.IO.Path]::GetExtension($fileName)

    Write-Host ""
    Write-Host "[$counter/$($filesToMove.Count)] Verwerken: $sourceUrl"

    # Standaard doel-URL
    $targetFileUrl = "$TargetFolderRelativeUrl/$fileName"

    # Check op duplicates in de doelmap
    $dupIndex = 1
    $finalFileName = $fileName

    while (Get-PnPFile -Url $targetFileUrl -ErrorAction SilentlyContinue) {
        $finalFileName = "{0}_DUP{1}{2}" -f $baseName, $dupIndex, $extension
        $targetFileUrl = "$TargetFolderRelativeUrl/$finalFileName"
        $dupIndex++
    }

    if ($WhatIf) {
        if ($finalFileName -ne $fileName) {
            Write-Host "  WOULD MOVE (DUPLICATE):" -ForegroundColor Yellow
            Write-Host "  From: $sourceUrl"
            Write-Host "  To:   $targetFileUrl"
        } else {
            Write-Host "  WOULD MOVE:" -ForegroundColor Yellow
            Write-Host "  From: $sourceUrl"
            Write-Host "  To:   $targetFileUrl"
        }
        continue
    }

    # 1) File verplaatsen
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
        Write-Host "  FOUT bij verplaatsen van $sourceUrl -> $targetFileUrl" -ForegroundColor Red
        Write-Host "  $_"
        continue
    }

    # 2) Metadata instellen om later te kunnen groeperen/mergen
    #    DuplicateGroupKey = originele bestandsnaam zonder extensie
    #    OriginalPath      = oorspronkelijke FileRef vóór verplaatsing
    Set-DuplicateMetaSafely -Item $item -DuplicateGroupKey $baseName -OriginalPath $sourceUrl
}

Write-Host ""
Write-Host "Klaar. Verwerkte bestanden: $counter"
