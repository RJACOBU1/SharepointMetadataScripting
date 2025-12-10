# ---------------------------------------------------------
# VERBINDEN MET JE SITE (pas aan naar jouw omgeving)
# ---------------------------------------------------------

# Connect-PnPOnline -Url "https://intranet.volvocars.net/sites/TestZone" `
#     -ClientId "10598e56-f5f6-42aa-bbcf-549965e26add" `
#     -Interactive

# ---------------------------------------------------------
# DECLARATIONS
# ---------------------------------------------------------

    # Naam van de documentbibliotheek
$listName = "Documents"

    # (server relative path na de site-URL, bv. /sites/TestZone/Shared Documents/Documents) of $null for everything.
$rootFolderFilter = "/sites/TestZone/Shared Documents/010 Crossconveyor"

    # Velden (interne namen, _x0020_ = spatie)
$col_EquipmentDocument   = "EquipmentDocument"     
$col_NameDocument        = "Name_x0020_Document"    
$col_StationLookup       = "StationNumbers"        
$col_MachineCodeLookup   = "MachineCode"            

    # Prefix voor station/machine codes (wordt vóór segment gezet, bv. "110CC102" -> "B-4122110CC102")
$machinePrefix = "B-4122"

# ---------------------------------------------------------
# Lookup-config voor StationNumbers (aparte lijst)
# ---------------------------------------------------------

$stationListName           = "Station_Numbers"
$stationLookupByNumber     = @{}      # StationNumber (tekst) -> ID
$stationCacheInitialized   = $false

function Initialize-StationCache {
    if ($stationCacheInitialized) { return }

    Write-Host "Station_Numbers cache initialiseren..." -ForegroundColor Cyan

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
# LOGICA:
    # Station + Machine info uit pad halen
    # Segmentvorm: 110CC102 / 120CR105 / 130CLT205 / ...
    # Layout van eindvorm: 3 cijfers + 2-3 letters + 3 cijfers
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
    # segments[0] = bv. "110CC102", "120CR105", "130CLT205", ...
    if ($segments.Length -lt 2) {
        return $null
    }

    $machineSegment = $segments[0]

    # Opdeling: 3 cijfers + 2-3 letters + 3 cijfers
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
    $pathLower    = $fileRef.ToLower()
    $fileNameLower = $fileName.ToLower()

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

        # Filename-regels binnen "electrical"
        if ($fileNameLower -like "*cable*") {
            $values[$col_NameDocument] = "2.1 Cable Calculation"
        }
        elseif ($fileNameLower -like "*ip address*") {
            $values[$col_NameDocument] = "2.2 IP addresses"
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
        $values[$col_NameDocument]      = "5.2 EC Declaration of Conformity"
    }

    if ($pathLower -like "*safety*") {
        $values[$col_EquipmentDocument] = "6. Safety"

        # Filename-regels binnen "safety"
        if ($fileNameLower -like "*inspection*" -and $fileNameLower -like "*report*") {
            $values[$col_NameDocument] = "6.2 Inspection reports"
        }
        elseif ($fileNameLower -like "*test*" -and $fileNameLower -like "*report*") {
            $values[$col_NameDocument] = "6.10 Safety test report"
        }
        elseif ($fileNameLower -like "*move*" -and $fileNameLower -like "*layout*") {
            $values[$col_NameDocument] = "6.6 Safe move layout"
        }
        elseif ($fileNameLower -like "*move*" -and $fileNameLower -like "*report*") {
            $values[$col_NameDocument] = "6.7 Safe move report"
        }
        elseif ($fileNameLower -like "*plc*" -and $fileNameLower -like "*report*") {
            $values[$col_NameDocument] = "6.12 Safety PLC report"
        }
        elseif ($fileNameLower -like "*matrix*") {
            $values[$col_NameDocument] = "6.9 Safety matrix"
        }
        elseif ($fileNameLower -like "*layout*") {
            $values[$col_NameDocument] = "6.8 Safety layout"
        }
        elseif ($fileNameLower -like "*placard*") {
            $values[$col_NameDocument] = "6.11 Safety placards"
        }
        elseif ($fileNameLower -like "*risk*" -or $fileNameLower -like "*assessment*") {
            $values[$col_NameDocument] = "6.5 Risk assessment"
        }
        elseif ($fileNameLower -like "*personal*" -or $fileNameLower -like "*equipment*") {
            $values[$col_NameDocument] = "6.3 Personal safety equipment"
        }
        elseif ($fileNameLower -like "*general*" -or $fileNameLower -like "*instruction*") {
            $values[$col_NameDocument] = "6.1 General safety instructions"
        }
    }

    if ($pathLower -like "*technical*") {
        $values[$col_EquipmentDocument] = "7. Design prerequisites"
        $values[$col_NameDocument]      = "7.3 Technical data"

        # Filename-regels binnen "technical"
        if ($fileNameLower -like "*fmea*") {
            $values[$col_NameDocument] = "7.1 FMEA"
        }
        elseif (($fileNameLower -like "*motor*" -or $fileNameLower -like "*belt*") -and $fileNameLower -like "*calc*") {
            $values[$col_NameDocument] = "7.4 Motor and belt calculation"
        }
        elseif ($fileNameLower -like "*stress*" -or $fileNameLower -like "*calculation*") {
            $values[$col_NameDocument] = "7.2 Stress calculation"
        }
        elseif ($fileNameLower -like "*time*" -or $fileNameLower -like "*analysis*") {
            $values[$col_NameDocument] = "7.6 Time analysis"
        }
    }

    if ($pathLower -like "*function*") {
        $values[$col_EquipmentDocument] = "8. Function Description"
        $values[$col_NameDocument]      = "8.1 Function description"

        # Filename-regels binnen "function"
        if ($fileNameLower -like "*tool*" -or $fileNameLower -like "*tree*") {
            $values[$col_NameDocument] = "8.2 Tool Tree"
        }
    }

    if ($pathLower -like "*operation*") {
        $values[$col_EquipmentDocument] = "9. Manuals"
        $values[$col_NameDocument]      = "9.1 Operation instructions"

        # Filename-regels binnen "operation"
        if ($fileNameLower -like "*training*" -or $fileNameLower -like "*education*") {
            $values[$col_NameDocument] = "9.2 Training material"
        }
    }

    if ($pathLower -like "*initial*") {
        $values[$col_EquipmentDocument] = "10. Settings & Measurements"
        $values[$col_NameDocument]      = "10.2 Initial settings data"

        # Filename-regels binnen "Initial"
        if ($fileNameLower -like "*energy*" -or $fileNameLower -like "*consumption*") {
            $values[$col_NameDocument] = "10.3 Energy consumption report"
        }
        elseif ($fileNameLower -like "*measurements*" -or $fileNameLower -like "*measuring*") {
            $values[$col_NameDocument] = "10.4 Measurements reports"
        }
        elseif ($fileNameLower -like "*clamp*" -or $fileNameLower -like "*force*") {
            $values[$col_NameDocument] = "10.1 Clamp/assembly force sheet"
        }
    }

    if ($pathLower -like "*maintenance*") {
        $values[$col_EquipmentDocument] = "11. Maintenance"
        $values[$col_NameDocument]      = "11.2 Maint-repair instructions"

        # Filename-regel binnen "maintenance"
        if ($fileNameLower -like "*matrix*") {
            $values[$col_NameDocument] = "11.1 Maintenance matrix"
        }
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
    #    (zelfde logica als je huidige script)
    # -----------------------------------------------------
    if ($col_StationLookup -or $col_MachineCodeLookup) {

        $stationInfo = Get-StationInfoFromPath -fileRef $fileRef

        if ($stationInfo -and $stationInfo.StationNumber) {

            # Station lookup
            if ($col_StationLookup) {
                $stationId = Get-StationLookupId -stationNumber $stationInfo.StationNumber
                if ($stationId) {
                    $values[$col_StationLookup] = $stationId
                }
            }

            # MachineCode lookup
            if ($col_MachineCodeLookup -and $stationInfo.MachineCode) {
                $machineId = Get-MachineLookupId -machineCode $stationInfo.MachineCode
                if ($machineId) {
                    $values[$col_MachineCodeLookup] = $machineId
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
