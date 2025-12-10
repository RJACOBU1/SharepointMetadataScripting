# ---------------------------------------------------------
# VERBINDEN MET JE SITE (pas aan naar jouw omgeving)
# ---------------------------------------------------------

# Import-Module PnP.PowerShell
# Connect-PnPOnline -Url "https://intranet.volvocars.net/sites/TestZone" `
#     -ClientId "10598e56-f5f6-42aa-bbcf-549965e26add" `
#     -Interactive

# ---------------------------------------------------------
# Instellingen
# ---------------------------------------------------------

# Naam van de documentbibliotheek
$listName = "Documents"              # bv. "Documents" of "Equipment Library"

# Optioneel: enkel items onder een bepaalde rootfolder verwerken
# (server relative path na de site-URL, bv. /sites/TestZone/Shared Documents/Documents)
# Zet op $null als je ALLES in de bibliotheek wil.
$rootFolderFilter = "/sites/TestZone/Shared Documents/010 Crossconveyor"
# $rootFolderFilter = "/sites/TestZone/Shared Documents/Documents"

# Velden (interne namen!)
$col_EquipmentDocument   = "EquipmentDocument"      # interne naam van "Equipment Document"
$col_NameDocument        = "Name_x0020_Document"    # interne naam van "Name Document"
$col_StationLookup       = "StationNumbers"         # interne naam van lookupkolom naar Station_Numbers
$col_MachineCodeLookup   = "MachineCode"            # interne naam van lookupkolom naar Maximo_Machine_Codes_List

# Prefix voor station/machine codes (wordt vóór segment gezet, bv. "110CC102" -> "B-4122110CC102")
$machinePrefix = "B-4122"

# ---------------------------------------------------------
# Lookup-config voor StationNumbers
# ---------------------------------------------------------

# ---------------------------------------------------------
# Lookup-config voor StationNumbers (aparte lijst)
# ---------------------------------------------------------

$stationListName           = "Station_Numbers"
$stationLookupByNumber     = @{}      # StationNumber (tekst) -> ID
$stationCacheInitialized   = $false

function Initialize-StationCache {
    if ($stationCacheInitialized) { return }

    Write-Host "Station_Numbers cache initialiseren..." -ForegroundColor Cyan

    # Haal alle items op met hun Title (waar StationNumber in zit)
    $items = Get-PnPListItem -List $stationListName -PageSize 2000 -Fields "Title"

    foreach ($it in $items) {
        $num = [string]$it["Title"]
        if ([string]::IsNullOrWhiteSpace($num)) { continue }

        $stationLookupByNumber[$num] = $it.Id
    }

    $global:stationCacheInitialized = $true
    Write-Host "Station_Numbers cache geladen: $($stationLookupByNumber.Count) stations." -ForegroundColor Green
}

function Get-StationLookupId {
    param(
        [string]$stationNumber
    )

    Initialize-StationCache

    if ($stationLookupByNumber.ContainsKey($stationNumber)) {
        return $stationLookupByNumber[$stationNumber]
    }

    Write-Warning "Geen Station-ID gevonden voor '$stationNumber' in lijst '$stationListName'."
    return $null
}


# ---------------------------------------------------------
# Lookup-config voor MachineCode (aparte lijst)
# ---------------------------------------------------------

$machineListName           = "Machine_Codes"
$machineLookupByCode       = @{}      # MachineCode (tekst) -> ID
$machineCacheInitialized   = $false

function Initialize-MachineCache {
    if ($machineCacheInitialized) { return }

    Write-Host "Machine_Codes cache initialiseren..." -ForegroundColor Cyan

    # We gaan ervan uit dat de MachineCode in Title staat
    $items = Get-PnPListItem -List $machineListName -PageSize 2000 -Fields "Title"

    foreach ($it in $items) {
        $code = [string]$it["Title"]
        if ([string]::IsNullOrWhiteSpace($code)) { continue }

        $machineLookupByCode[$code] = $it.Id
    }

    $global:machineCacheInitialized = $true
    Write-Host "Machine_Codes cache geladen: $($machineLookupByCode.Count) codes." -ForegroundColor Green
}

function Get-MachineLookupId {
    param(
        [string]$machineCode
    )

    Initialize-MachineCache

    if ($machineLookupByCode.ContainsKey($machineCode)) {
        return $machineLookupByCode[$machineCode]
    }

    Write-Warning "Geen Machine-ID gevonden voor '$machineCode' in lijst '$machineListName'."
    return $null
}


# ---------------------------------------------------------
# Station + Machine info uit pad halen
# Segmentvorm: 110CC102 / 120CR105 / 130CLT205 / ...
# Layout: 3 cijfers + 2-3 letters + 3 cijfers
# StationNumber = prefix + eerste 3 cijfers   -> B-4122110
# MachineCode   = prefix + volledige segment  -> B-4122110CC102
# ---------------------------------------------------------

function Get-StationInfoFromPath {
    param(
        [string]$fileRef   # volledige server-relative URL
    )

    # Maak pad eventueel relatief t.o.v. rootFolderFilter
    $relativePath = $fileRef
    if ($rootFolderFilter -and
        $fileRef.StartsWith($rootFolderFilter, [System.StringComparison]::OrdinalIgnoreCase)) {

        $relativePath = $fileRef.Substring($rootFolderFilter.Length).Trim('/','\')
    }

    # Split in segmenten
    $segments = $relativePath -split '[\\/]+'

    Write-Host "  [station-scan] Relatief pad: $relativePath" -ForegroundColor DarkGray
    Write-Host "  [station-scan] Segments: $($segments -join ' | ')" -ForegroundColor DarkGray

    # We verwachten:
    # segments[1] = bv. "110CC102", "120CR105", "130CLT205", ...
    if ($segments.Length -lt 2) {
        return $null
    }

    $machineSegment = $segments[0]

    # Patroon: 3 cijfers + 2-3 letters + 3 cijfers
    if ($machineSegment -match '^(?<base>\d{3})(?<letters>[A-Za-z]{2,3})(?<digits>\d{3})$') {

        $base        = $matches['base']   # bv. "110"
        $stationNumber = "$machinePrefix$base"          # B-4122 + 110  -> B-4122110
        $machineCode   = "$machinePrefix$machineSegment" # B-4122 + 110CC102 -> B-4122110CC102

        Write-Host "  [station] StationNumber = '$stationNumber'" -ForegroundColor Cyan
        Write-Host "  [station] MachineCode   = '$machineCode'"   -ForegroundColor Cyan

        return [PSCustomObject]@{
            StationNumber = $stationNumber
            MachineCode   = $machineCode
        }
    }
    else {
        Write-Host "  [station] Segment '$machineSegment' matcht patroon niet (3 cijfers + 2-3 letters + 3 cijfers)." -ForegroundColor DarkYellow
        return $null
    }
}

# ---------------------------------------------------------
# Functie: bepaal metadata voor één file op basis van pad + extensie + station/machine info
# ---------------------------------------------------------
function Get-MetadataForFile {
    param(
        [string]$fileRef,    # volledige server-relative URL
        [string]$fileName,   # enkel bestandsnaam
        [string]$extension   # extensie met punt, in lower case
    )

    # Hashtable met updates die we willen zetten
    $values = @{}

    # Alles case-insensitive maken
    $pathLower = $fileRef.ToLower()

    # -----------------------------------------------------
    # 1. Discipline-regels (mechanical / electrical / pneumatic)
    # -----------------------------------------------------

    if ($pathLower -like "*mechanical*") {
        $values[$col_EquipmentDocument] = "1. Mechanical"

        # - pdf / dwg in mechanical → 1.1 Mechanical Drawing
        # - xls in mechanical       → 1.2 Mechanical Part list
        if (-not $values.ContainsKey($col_NameDocument)) {
            if ($extension -in @(".pdf", ".dwg")) {
                $values[$col_NameDocument] = "1.1 Mechanical Drawing"
            }
            elseif ($extension -in @(".xls", ".xlsx")) {
                $values[$col_NameDocument] = "1.2 Mechanical Part list"
            }
        }
    }
    elseif ($pathLower -like "*electrical*") {
        $values[$col_EquipmentDocument] = "2. Electrical"

        # - xls in electrical → 2.4 Electrical Part list
        if (-not $values.ContainsKey($col_NameDocument)) {
            if ($extension -in @(".xls", ".xlsx")) {
                $values[$col_NameDocument] = "2.4 Electrical Part list"
            }
        }
    }
    elseif ($pathLower -like "*pneumatic*") {
        $values[$col_EquipmentDocument] = "3. Pneumatic/Hydraulic/Media"

        # - xls in pneumatic → 3.2 Pneumatic/Hydraulic/Media Part list
        if (-not $values.ContainsKey($col_NameDocument)) {
            if ($extension -in @(".xls", ".xlsx")) {
                $values[$col_NameDocument] = "3.2 Pneumatic/Hydraulic/Media Part list"
            }
        }
    }

    # -----------------------------------------------------
    # 2. Documenttype-regels (overschrijven disciplines indien nodig)
    # -----------------------------------------------------

    if ($pathLower -like "*datasheet*") {
        $values[$col_EquipmentDocument] = "4. Data Sheet"
        $values[$col_NameDocument]      = "4.1 Data Sheet"
    }

    if ($pathLower -like "*legal*") {
        $values[$col_EquipmentDocument] = "5. Legal"
    }

    if ($pathLower -like "*safety*") {
        $values[$col_EquipmentDocument] = "6. Safety"
    }

    if ($pathLower -like "*technical*") {
        $values[$col_EquipmentDocument] = "7. Design prerequisites"
        $values[$col_NameDocument]      = "7.3 Technical data"
    }

    if ($pathLower -like "*function*") {
        $values[$col_EquipmentDocument] = "8. Function Description"
        $values[$col_NameDocument]      = "8.1 Function description"
    }

    if ($pathLower -like "*operation*") {
        $values[$col_EquipmentDocument] = "9. Manuals"
        $values[$col_NameDocument]      = "9.1 Operation instructions"
    }

    if ($pathLower -like "*initial*") {
        $values[$col_EquipmentDocument] = "10. Settings & Measurements"
        $values[$col_NameDocument]      = "10.2 Initial settings data"
    }

    if ($pathLower -like "*measuring*") {
        $values[$col_EquipmentDocument] = "10. Settings & Measurements"
        $values[$col_NameDocument]      = "10.4 Measurements reports"
    }

    if ($pathLower -like "*maintenance*") {
        $values[$col_EquipmentDocument] = "11. Maintenance"
        $values[$col_NameDocument]      = "11.2 Maint-repair instructions"
    }

    if ($pathLower -like "*service*") {
        $values[$col_EquipmentDocument] = "12. Supplier information"
        $values[$col_NameDocument]      = "12.1 Warranty & service agreement"
    }

    # -----------------------------------------------------
    # 3. Extensie-specifieke override: mp4 = manuals
    # -----------------------------------------------------
    if ($extension -eq ".mp4") {
        $values[$col_EquipmentDocument] = "9. Manuals"
        $values[$col_NameDocument]      = "9.1 Operation instructions"
    }

    # -----------------------------------------------------
    # 4. Stationnummer + MachineCode uit pad halen + beide lookups zetten
    # -----------------------------------------------------
    if ($col_StationLookup -or $col_MachineCodeLookup) {

        $stationInfo = Get-StationInfoFromPath -fileRef $fileRef

        if ($stationInfo -and $stationInfo.StationNumber) {

            # Station lookup
            if ($col_StationLookup) {
                $stationId = Get-StationLookupId -stationNumber $stationInfo.StationNumber
                if ($stationId) {
                    Write-Host "  Station gevonden: $($stationInfo.StationNumber) → Station ID = $stationId" -ForegroundColor Cyan
                    $values[$col_StationLookup] = $stationId
                }
                else {
                    Write-Host "  [station] Geen Station-ID gevonden voor '$($stationInfo.StationNumber)'." -ForegroundColor DarkYellow
                }
            }

            # MachineCode lookup
            if ($col_MachineCodeLookup -and $stationInfo.MachineCode) {
                $machineId = Get-MachineLookupId -machineCode $stationInfo.MachineCode
                if ($machineId) {
                    Write-Host "  MachineCode gevonden: $($stationInfo.MachineCode) → Machine ID = $machineId" -ForegroundColor Cyan
                    $values[$col_MachineCodeLookup] = $machineId
                }
                else {
                    Write-Host "  [machine] Geen Machine-ID gevonden voor '$($stationInfo.MachineCode)'." -ForegroundColor DarkYellow
                }
            }
        }
    }

    return $values
}

# ---------------------------------------------------------
# Hoofdscript: loop over alle items en pas regels toe
# ---------------------------------------------------------

Write-Host "Ophalen van items uit lijst '$listName'..." -ForegroundColor Cyan

Get-PnPListItem -List $listName -PageSize 2000 -ScriptBlock {
    param($items)

    foreach ($item in $items) {

        # Enkel files, geen mappen
        if ($item.FileSystemObjectType -ne "File") {
            continue
        }

        $fileRef  = $item["FileRef"]      # volledige server relative URL
        $fileName = $item["FileLeafRef"]  # enkel bestandsnaam

        # Filter op rootFolder indien ingesteld
        if ($rootFolderFilter -and ($fileRef -notlike "$rootFolderFilter*")) {
            continue
        }

        # Extensie bepalen (lowercase)
        $extension = [System.IO.Path]::GetExtension($fileName).ToLower()

        # Metadata bepalen via onze functie
        $valuesToSet = Get-MetadataForFile -fileRef $fileRef -fileName $fileName -extension $extension

        # Als er niets te zetten is, sla over
        if (-not $valuesToSet -or $valuesToSet.Count -eq 0) {
            continue
        }

        Write-Host "Update item $($item.Id): $fileName" -ForegroundColor Yellow
        $valuesToSet.GetEnumerator() | ForEach-Object {
            Write-Host "  $($_.Key) = $($_.Value)"
        }

        # Metadata wegschrijven naar SharePoint
        Set-PnPListItem -List $listName -Identity $item.Id -Values $valuesToSet | Out-Null
    }
}

Write-Host "Klaar. Alle metadata-regels + station- en machine-lookups toegepast." -ForegroundColor Green
