# Document library
$lib = Get-PnPList -Identity "Documents"

# Rootfolder waaronder je structuur zit:
$rootFolder = "Shared Documents/Documents"

# Lookup-lijst voor StationNumbers
$stationListName = "Station_Numbers"

# Cache om lookup-ID's niet telkens opnieuw te moeten opvragen
$stationLookupCache = @{}


# functie: lookup-ID ophalen (en cachen) op basis van stationnummer
function Get-StationLookupId {
    param(
        [string]$stationNumber
    )

    if ($stationLookupCache.ContainsKey($stationNumber)) {
        return $stationLookupCache[$stationNumber]
    }

    $lookupItem = Get-PnPListItem -List $stationListName -Query @"
<View>
  <Query>
    <Where>
      <Eq>
        <FieldRef Name='Title' />
        <Value Type='Text'>$stationNumber</Value>
      </Eq>
    </Where>
  </Query>
</View>
"@

    if (-not $lookupItem) {
        Write-Warning "Geen lookup gevonden voor station '$stationNumber' in lijst '$stationListName'."
        return $null
    }

    $id = $lookupItem.Id
    $stationLookupCache[$stationNumber] = $id
    return $id
}


# -------------------------------------------------------
# 1) Alle "equipment"-subfolders onder root ophalen
#    vb. "1 Mechanical", "2 Electrical", ...
# -------------------------------------------------------

$equipmentFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $rootFolder -ItemType Folder

foreach ($eqFolder in $equipmentFolders) {

    $equipmentDocument = $eqFolder.Name        # bv. "1 Mechanical"

    # EquipmentDocument afleiden:
    # "1 Mechanical" -> "1. Mechanical"
    $equipmentDocument = $equipmentFolderName -replace '^(\d+)\s+', '$1. '

    Write-Host "Verwerk equipment-folder: $equipmentFolderName (EquipmentDocument = '$equipmentDocument')" -ForegroundColor Cyan

    # ---------------------------------------------------
    # 2) Alle station-subfolders onder deze equipment-map
    #    vb. "B-121210", "B-121211", ...
    # ---------------------------------------------------

    $equipmentFolderPath = "$rootFolder/$equipmentFolderName"

    $stationFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $equipmentFolderPath -ItemType Folder

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
            $li = Get-PnPFile -Url $f.ServerRelativeUrl -AsListItem

            Set-PnPListItem -List $lib -Identity $li.Id -Values @{
                'EquipmentDocument' = $equipmentDocument
                'Manufacturer'      = 'Volvo'
                'StationNumbers'    = $lookupId
            } -UpdateType Update
        }
    }
}
