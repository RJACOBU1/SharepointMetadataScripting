# ---------------------------------------------------------
# VERBINDEN MET JE SITE (pas aan naar jouw omgeving)
# ---------------------------------------------------------

# Connect-PnPOnline -Url "https://intranet.volvocars.net/sites/TestZone" `
#     -ClientId "10598e56-f5f6-42aa-bbcf-549965e26add" `
#     -Interactive

# ---------------------------------------------------------
# Instellingen
# ---------------------------------------------------------

# Naam van de documentbibliotheek
$listName = "Documents"              # bv. "Documents" of "Equipment Library"

# Optioneel: enkel items onder een bepaalde rootfolder verwerken
# (server relative path na de site-URL, bv. /sites/TestZone/Shared Documents/GEICO)
# Zet op $null als je ALLES in de bibliotheek wil.
$rootFolderFilter = "/sites/TestZone/Shared Documents/010 Crossconveyor"
# $rootFolderFilter = "/sites/TestZone/Shared Documents/GEICO"

# Velden (interne namen!)
$col_EquipmentDocument = "EquipmentDocument"      # interne naam van "Equipment Document"
$col_NameDocument      = "Name_x0020_Document"    # interne naam van "Name Document"
$col_StationLookup     = "StationNumbers"         # interne naam van lookupkolom naar Station_Numbers

# ---------------------------------------------------------
# Stationnummer lookup-config
# ---------------------------------------------------------

# Lookup-lijst voor StationNumbers
$stationListName = "Station_Numbers"

# Cache: stationnummer -> lookup-ID
$stationLookupCache = @{}
# Extra cache met alleen de stationnummers (titles) voor snelle path-matching
$stationTitleCache = @()

function Initialize-StationCache {
    if ($stationTitleCache.Count -gt 0) {
        return
    }

    Write-Host "Station_Numbers cache initialiseren..." -ForegroundColor Cyan

    $items = Get-PnPListItem -List $stationListName -PageSize 2000 -Fields "Title"

    foreach ($it in $items) {
        $rawTitle = [string]$it["Title"]
        $title    = $rawTitle.Trim()    # heel belangrijk: spaties weg!

        if ([string]::IsNullOrWhiteSpace($title)) { continue }

        # Vul beide caches
        $stationLookupCache[$title] = $it.Id
        $stationTitleCache += $title
    }

    Write-Host "Station_Numbers cache geladen: $($stationTitleCache.count) stations." -ForegroundColor Green
}

# functie: lookup-ID ophalen (en cachen) op basis van stationnummer
function Get-StationLookupId {
    param(
        [string]$stationNumber
    )

    Initialize-StationCache

    if ($stationLookupCache.ContainsKey($stationNumber)) {
        return $stationLookupCache[$stationNumber]
    }

    # Indien niet gevonden in cache: val terug op CAML-query (zou uitzonderlijk moeten zijn)
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

# Stationnummer detecteren in het file pad
function Get-StationNumberFromPath {
    param(
        [string]$fileRef
    )

    Initialize-StationCache

    $pathLower = $fileRef.ToLower()

    # DEBUG: toon eenmaal het pad (eerste paar keer)
    Write-Host "  [station-scan] Pad: $fileRef" -ForegroundColor DarkGray

    foreach ($title in $stationTitleCache) {
        $titleTrim  = $title.Trim()
        $titleLower = $titleTrim.ToLower()

        if ($pathLower.Contains($titleLower)) {
            Write-Host "  [station-scan] Match gevonden: '$titleTrim' in pad." -ForegroundColor DarkCyan
            return $titleTrim
        }
    }

    return $null
}


# ---------------------------------------------------------
# Functie: bepaal metadata voor één file op basis van pad + extensie + stationnummer
# ---------------------------------------------------------
function Get-MetadataForFile {
    param(
        [string]$fileRef,    # volledige server relative URL (bv. /sites/.../Shared Documents/Mechanical/...)
        [string]$fileName,   # enkel bestandsnaam (bv. Pump_01.dwg)
        [string]$extension   # extensie met punt, in lower case (bv. .pdf)
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

        # Extra regels:
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

        # Extra regels:
        # - xls in electrical → 2.4 Electrical Part list
        if (-not $values.ContainsKey($col_NameDocument)) {
            if ($extension -in @(".xls", ".xlsx")) {
                $values[$col_NameDocument] = "2.4 Electrical Part list"
            }
        }
    }
    elseif ($pathLower -like "*pneumatic*") {
        $values[$col_EquipmentDocument] = "3. Pneumatic/Hydraulic/Media"

        # Extra regels:
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

    # Datasheet
    if ($pathLower -like "*datasheet*") {
        $values[$col_EquipmentDocument] = "4. Data Sheet"
        $values[$col_NameDocument]      = "4.1 Data Sheet"
    }

    # Legal
    if ($pathLower -like "*legal*") {
        $values[$col_EquipmentDocument] = "5. Legal"
    }

    # Safety
    if ($pathLower -like "*safety*") {
        $values[$col_EquipmentDocument] = "6. Safety"
    }

    # Technical
    if ($pathLower -like "*technical*") {
        $values[$col_EquipmentDocument] = "7. Design prerequisites"
        $values[$col_NameDocument]      = "7.3 Technical data"
    }

    # Function
    if ($pathLower -like "*function*") {
        $values[$col_EquipmentDocument] = "8. Function Description"
        $values[$col_NameDocument]      = "8.1 Function description"
    }

    # Operation
    if ($pathLower -like "*operation*") {
        $values[$col_EquipmentDocument] = "9. Manuals"
        $values[$col_NameDocument]      = "9.1 Operation instructions"
    }

    # Initial
    if ($pathLower -like "*initial*") {
        $values[$col_EquipmentDocument] = "10. Settings & Measurements"
        $values[$col_NameDocument]      = "10.2 Initial settings data"
    }

    # Measuring
    if ($pathLower -like "*measuring*") {
        $values[$col_EquipmentDocument] = "10. Settings & Measurements"
        $values[$col_NameDocument]      = "10.4 Measurements reports"
    }

    # Maintenance
    if ($pathLower -like "*maintenance*") {
        $values[$col_EquipmentDocument] = "11. Maintenance"
        $values[$col_NameDocument]      = "11.2 Maint-repair instructions"
    }

    # Service
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
    # 4. Stationnummer uit pad halen en lookup invullen
    # -----------------------------------------------------
    if ($col_StationLookup) {
        $stationNumber = Get-StationNumberFromPath -fileRef $fileRef
        if ($stationNumber) {
            $stationId = Get-StationLookupId -stationNumber $stationNumber
            if ($stationId) {
                Write-Host "  Station gevonden in pad: $stationNumber (ID = $stationId)" -ForegroundColor Cyan

                # Single-value lookup: FieldLookupValue gebruiken
                $values[$col_StationLookup] = $stationId
            }
        }
    }
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

Write-Host "Klaar. Alle metadata-regels + stationlookup toegepast." -ForegroundColor Green
