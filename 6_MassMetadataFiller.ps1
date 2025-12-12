# =============== CONFIG ===============

$listName = "Documents"

# Enkel items onder deze folder (zelfde stijl als je andere script)
$rootFolderFilter = "/sites/TestZone/Shared Documents/FLATTENED"

# Kolom + vaste waarde
$fieldInternalName = "Area"   # <-- interne naam (vaak is dit gewoon "Area")
$fixedChoiceValue  = "Pre-Treatment Preparation"

# Performance
$pageSize = 2000
# =====================================


Write-Host "Ophalen items uit '$listName' onder '$rootFolderFilter'..." -ForegroundColor Cyan

# Haal enkel bestanden op (FSObjType=0), onder je rootfolder, en waar Area nog niet gelijk is aan de gewenste waarde
#HIER PAS JE Value Type aan naargelang --> 'Choice', 'Text', 
$caml = @"
<View Scope='RecursiveAll'>
  <Query>
    <Where>
      <And>
        <!-- Folder filter -->
        <BeginsWith>
          <FieldRef Name='FileDirRef' />
          <Value Type='Text'>$rootFolderFilter</Value>
        </BeginsWith>

        <And>
          <!-- Only files -->
          <Eq>
            <FieldRef Name='FSObjType'/>
            <Value Type='Integer'>0</Value>
          </Eq>

          <!-- Only if not already correct -->
          <Or>
            <IsNull>
              <FieldRef Name='$fieldInternalName'/>
            </IsNull>
            <Neq>
              <FieldRef Name='$fieldInternalName'/>
              <Value Type='Text'>$fixedChoiceValue</Value>
            </Neq>
          </Or>
        </And>
      </And>
    </Where>
  </Query>

  <ViewFields>
    <FieldRef Name='ID'/>
    <FieldRef Name='FileRef'/>
    <FieldRef Name='FileDirRef'/>
    <FieldRef Name='$fieldInternalName'/>
  </ViewFields>

  <RowLimit Paged='TRUE'>$pageSize</RowLimit>
</View>
"@


$items = Get-PnPListItem -List $listName -Query $caml

Write-Host "Aantal items om te updaten: $($items.Count)" -ForegroundColor Yellow
if ($items.Count -eq 0) { Write-Host "Niets te doen."; return }

# Batch (sneller dan één-per-één)
$batch = New-PnPBatch
$cnt = 0

foreach ($it in $items) {
    $cnt++

    # Zet vaste choice waarde
    Set-PnPListItem -List $listName -Identity $it.Id -Values @{
        $fieldInternalName = $fixedChoiceValue
    } -Batch $batch

    # Commit per 200 updates (veilig & snel)
    if (($cnt % 200) -eq 0) {
        Invoke-PnPBatch -Batch $batch
        $batch = New-PnPBatch
        Write-Host "Gecommit: $cnt / $($items.Count)"
    }
}

# laatste rest
Invoke-PnPBatch -Batch $batch
Write-Host "Klaar. Totaal geüpdatet: $cnt" -ForegroundColor Green
