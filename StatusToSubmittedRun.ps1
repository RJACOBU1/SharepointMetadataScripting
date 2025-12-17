# ================= CONFIG =================
$LibraryName = "Documents"
$RootFolderServerRelativeUrl = "/sites/TestZone/Shared Documents/B-(machine number)/100 SpecDoc/02 Technical Data"
$StatusFieldInternalName = "Status"
$NewStatusValue = "Submitted"

# ================= CAML QUERY =================
$camlQuery = @"
<View Scope='RecursiveAll'>
  <Query>
    <Where>
      <And>
        <!-- Enkel bestanden -->
        <Eq>
          <FieldRef Name='FSObjType' />
          <Value Type='Integer'>0</Value>
        </Eq>

        <!-- Status is leeg -->
        <IsNull>
          <FieldRef Name='$StatusFieldInternalName' />
        </IsNull>
      </And>
    </Where>
  </Query>
</View>
"@

# ================= GET ITEMS =================
$items = Get-PnPListItem `
    -List $LibraryName `
    -FolderServerRelativeUrl $RootFolderServerRelativeUrl `
    -Query $camlQuery `
    -PageSize 200

Write-Host "Gevonden bestanden: $($items.Count)"

# ================= UPDATE STATUS =================
foreach ($item in $items) {
    Set-PnPListItem `
        -List $LibraryName `
        -Identity $item.Id `
        -Values @{
            $StatusFieldInternalName = $NewStatusValue
        }

    Write-Host "Updated file ID $($item.Id) -> Status = Submitted"
}

Write-Host "Klaar."
