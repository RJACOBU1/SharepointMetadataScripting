# -------------------------------------------------------
# 1) Alle "equipment"-subfolders onder root ophalen
#    vb. "1 Mechanical", "2 Electrical", "3 Datasheet", ...
# -------------------------------------------------------

$equipmentFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $rootFolder -ItemType Folder

foreach ($eqFolder in $equipmentFolders) {

    $equipmentFolderName = $eqFolder.Name        # bv. "1 Mechanical", "3 Datasheet"

    # EquipmentDocument afleiden
    switch -regex ($equipmentFolderName) {
        '^1\s+Mechanical' { $equipmentDocument = '1. Mechanical'; break }
        '^2\s+Electrical' { $equipmentDocument = '2. Electrical'; break }
        '^3\s+Datasheet'  { $equipmentDocument = '4. Datasheet'; break }  # speciale mapping
        default {
            # fallback: "X Something" -> "X. Something"
            $equipmentDocument = [regex]::Replace($equipmentFolderName, '^(\d+)\s+', '$1. ')
        }
    }

    Write-Host "Verwerk equipment-folder: $equipmentFolderName (EquipmentDocument = '$equipmentDocument')" -ForegroundColor Cyan

    $equipmentFolderPath = "$rootFolder/$equipmentFolderName"

    # ---------------------------------------------------
    # 2) Alle station-subfolders onder deze equipment-map
    #    vb. "B-121210", "B-121211", ...
    # ---------------------------------------------------

    $stationFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $equipmentFolderPath -ItemType Folder

    if ($stationFolders.Count -eq 0 -and $equipmentDocument -eq '4. Datasheet') {
        # CASE: 3 Datasheet heeft geen subfolders → files zitten rechtstreeks in deze map

        Write-Host "  Geen station-subfolders, verwerk datasheets rechtstreeks in $equipmentFolderPath" -ForegroundColor Magenta

        $files = Get-PnPFolderItem -FolderSiteRelativeUrl $equipmentFolderPath -ItemType File

        foreach ($f in $files) {

            # Manufacturer uit bestandsnaam halen
            $manufacturer = $null
            $fileName = $f.Name   # vb. D23-069937836-01-001-Honeywell-V5049A1607_datasheet_EN

            $parts = $fileName.Split('-')
            if ($parts.Count -ge 5) {
                $manufacturer = $parts[4]
            } else {
                Write-Warning "    Kon merk niet afleiden uit bestandsnaam: $fileName"
            }

            $li = Get-PnPFile -Url $f.ServerRelativeUrl -AsListItem

            Set-PnPListItem -List $lib -Identity $li.Id -Values @{
                'EquipmentDocument' = $equipmentDocument   # 4. Datasheet
                'Manufacturer'      = $manufacturer        # bv. Honeywell
                'StationNumbers'    = $null                # of gewoon weglaten uit de hashtable als je echt niets wil zetten
            } -UpdateType Update
        }

        # Ga naar de volgende equipment-folder
        continue
    }

    # Als er wél station-subfolders zijn → oude gedrag
    foreach ($stFolder in $stationFolders) {

        $stationNumber = $stFolder.Name   # bv. "B-121210"
        Write-Host "  Verwerk station-folder: $stationNumber" -ForegroundColor Yellow

        # Lookup-ID ophalen voor dit stationnummer
        $lookupId = Get-StationLookupId -stationNumber $stationNumber

        if (-not $lookupId) {
            Write-Warning "  -> Station '$stationNumber' wordt overgeslagen (geen lookup-ID)."
            continue
        }

        # Pad naar deze stationmap
        $stationFolderPath = "$equipmentFolderPath/$stationNumber"

        # ---------------------------------------------------
        # 3) Alle files in deze stationmap ophalen
        # ---------------------------------------------------
        $files = Get-PnPFolderItem -FolderSiteRelativeUrl $stationFolderPath -ItemType File

        foreach ($f in $files) {

            # Default: Manufacturer leeg
            $manufacturer = $null

            # Alleen bij datasheets manufacturer invullen
            if ($equipmentDocument -eq '4. Datasheet') {
                $fileName = $f.Name
                $parts = $fileName.Split('-')
                if ($parts.Count -ge 5) {
                    $manufacturer = $parts[4]
                } else {
                    Write-Warning "    Kon merk niet afleiden uit bestandsnaam: $fileName"
                    
                }
            }

            $li = Get-PnPFile -Url $f.ServerRelativeUrl -AsListItem

            Set-PnPListItem -List $lib -Identity $li.Id -Values @{
                'EquipmentDocument' = $equipmentDocument
                'Manufacturer'      = $manufacturer
                'StationNumbers'    = $lookupId
            } -UpdateType Update
        }
    }
}
