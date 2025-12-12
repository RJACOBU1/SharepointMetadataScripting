param(
    [string]$SiteUrl = "https://intranet.volvocars.net/sites/TestZone",
    [string]$LibraryName = "Documents",
    [string]$FlattenedFolderRelativeUrl = "/sites/TestZone/Shared Documents/FLATTENED",
    [switch]$WhatIf
)


# Velden die we willen mergen (interne namen)
$lookupFields = @(
    "EquipmentDocument",
    "Name_x0020_Document",
    "StationNumberEff",
    "MachineCodeEff"
)

# Field metadata ophalen (bestaat? multi?)
$fieldInfo = @{}  # naam -> @{ Exists=bool; Multi=bool; Type=string }
$validLookupFields = @()

foreach ($fname in $lookupFields) {
    $f = Get-PnPField -List $LibraryName -Identity $fname -ErrorAction SilentlyContinue
    if (-not $f) {
        Write-Host "WAARSCHUWING: veld '$fname' bestaat niet en wordt genegeerd." -ForegroundColor Yellow
        continue
    }

    # Detecteer multi-lookup
    $isMulti = $false
    try { $isMulti = [bool]$f.AllowMultipleValues } catch { $isMulti = $false }

    $fieldInfo[$fname] = @{
        Exists = $true
        Multi  = $isMulti
        Type   = $f.TypeAsString
    }

    $validLookupFields += $fname
}

if ($validLookupFields.Count -eq 0) {
    Write-Host "Geen geldige velden gevonden. Stop."
    return
}

Write-Host "Velden die gemerged worden:"
foreach ($n in $validLookupFields) {
    Write-Host (" - {0} (Multi={1}, Type={2})" -f $n, $fieldInfo[$n].Multi, $fieldInfo[$n].Type)
}

# Items ophalen enkel uit FLATTENED
$fieldsToLoad = @("FileLeafRef","FileRef","FSObjType","ID") + $validLookupFields

Write-Host "Ophalen van items onder folder: $FlattenedFolderRelativeUrl ..."
$items = Get-PnPListItem -List $LibraryName -PageSize 500 `
    -FolderServerRelativeUrl $FlattenedFolderRelativeUrl `
    -Fields $fieldsToLoad `
    -ScriptBlock { param($b) $b }

$files = $items | Where-Object { $_["FSObjType"] -eq 0 }

Write-Host "Aantal bestanden in FLATTENED: $($files.Count)"
if ($files.Count -eq 0) { return }

function Get-LookupIds {
    param($lv)
    $ids = @()
    if (-not $lv) { return $ids }

    if ($lv -is [System.Array]) {
        foreach ($x in $lv) {
            if ($x -and $x.LookupId) { $ids += [int]$x.LookupId }
        }
    } else {
        if ($lv.LookupId) { $ids += [int]$lv.LookupId }
    }
    return $ids
}

# MergeKey op basis van bestandsnaam zonder _DUPx
foreach ($item in $files) {
    $fileName = $item["FileLeafRef"]
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)

    $mergeKey = if ($baseName -match "^(.*)_DUP\d+$") { $Matches[1] } else { $baseName }
    $item | Add-Member -NotePropertyName "MergeKey" -NotePropertyValue $mergeKey -Force
}

$groups = $files | Group-Object -Property MergeKey
Write-Host "Aantal mogelijke groepen: $($groups.Count)"

$groupsProcessed = 0
$totalDuplicatesRemoved = 0

foreach ($group in $groups) {
    $groupItems = $group.Group
    if ($groupItems.Count -le 1) { continue }

    $groupsProcessed++
    Write-Host ""
    Write-Host "=== Groep '$($group.Name)' ($($groupItems.Count) files) ==="

    # Master = zonder _DUP, anders laagste ID
    $master = $groupItems | Where-Object {
        ([System.IO.Path]::GetFileNameWithoutExtension($_["FileLeafRef"])) -notmatch "_DUP\d+$"
    } | Select-Object -First 1

    if (-not $master) {
        $master = $groupItems | Sort-Object { $_["ID"] } | Select-Object -First 1
        Write-Host "Geen niet-DUP gevonden; master = laagste ID: $($master['FileRef'])" -ForegroundColor Yellow
    } else {
        Write-Host "Master: $($master['FileRef'])"
    }

    $duplicates = $groupItems | Where-Object { $_ -ne $master }
    if ($duplicates.Count -eq 0) { continue }

    $updateValues = @{}

    foreach ($fieldName in $validLookupFields) {
        $isMulti = $fieldInfo[$fieldName].Multi

        $allIds = @()
        $allIds += Get-LookupIds $master[$fieldName]
        foreach ($dup in $duplicates) {
            $allIds += Get-LookupIds $dup[$fieldName]
        }

        $allIds = $allIds | Select-Object -Unique

        if ($allIds.Count -eq 0) { continue }

        if ($isMulti) {
            # MULTI LOOKUP: schrijf alle IDs
            # (Set-PnPListItem verwacht hiervoor een int[] / array)
            $updateValues[$fieldName] = @($allIds)
            Write-Host "  Merge MULTI lookup $fieldName -> IDs: $($allIds -join ', ')"
        }
        else {
            # SINGLE LOOKUP: behoud master als hij al iets heeft, anders neem eerste
            $masterIds = Get-LookupIds $master[$fieldName]
            $targetId = $null

            if ($masterIds.Count -gt 0) {
                $targetId = $masterIds[0]
            } else {
                $targetId = $allIds[0]
            }

            # Alleen zetten als master leeg was of anders
            if ($masterIds.Count -eq 0 -or $masterIds[0] -ne $targetId) {
                $updateValues[$fieldName] = [int]$targetId
                Write-Host "  Merge SINGLE lookup $fieldName -> ID: $targetId"
            }

            if ($allIds.Count -gt 1) {
                Write-Host "  LET OP: meerdere IDs gevonden voor SINGLE lookup '$fieldName' (groep '$($group.Name)')." -ForegroundColor Yellow
            }
        }
    }

    # Update master
    if ($updateValues.Count -gt 0) {
        if ($WhatIf) {
            Write-Host "  (WhatIf) zou master updaten: ID $($master['ID'])" -ForegroundColor Yellow
        } else {
            try {
                Set-PnPListItem -List $LibraryName -Identity $master["ID"] -Values $updateValues -ErrorAction Stop | Out-Null
                Write-Host "  Master geüpdatet (ID $($master['ID']))"
            } catch {
                Write-Host "  FOUT bij updaten master ID $($master['ID']) :" -ForegroundColor Red
                Write-Host $_
                continue
            }
        }
    } else {
        Write-Host "  Geen metadata changes nodig."
    }

    # Duplicates verwijderen
    foreach ($dup in $duplicates) {
        $dupUrl = $dup["FileRef"]
        if ($WhatIf) {
            Write-Host "  (WhatIf) zou verwijderen: $dupUrl" -ForegroundColor Yellow
        } else {
            try {
                Remove-PnPFile -ServerRelativeUrl $dupUrl -Recycle -Force -ErrorAction Stop
                Write-Host "  Verwijderd (recycle): $dupUrl"
                $totalDuplicatesRemoved++
            } catch {
                Write-Host "  FOUT bij verwijderen $dupUrl :" -ForegroundColor Red
                Write-Host $_
            }
        }
    }
}

Write-Host ""
Write-Host "=== SAMENVATTING ==="
Write-Host "Groepen verwerkt: $groupsProcessed"
Write-Host "Duplicates verwijderd: $totalDuplicatesRemoved"
