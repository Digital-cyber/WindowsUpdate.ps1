# WindowsUpdate.ps1

$session = New-Object -ComObject Microsoft.Update.Session 
$session.ClientApplicationID = "PSWindowsUpdate"

$searcher = $session.CreateUpdateSearcher()
$updates = $searcher.Search("IsInstalled=0").Updates

$installer = $session.CreateUpdateInstaller()
$installer.Updates = $updates 
$installer.ForceQuiet = $true  

Write-Host "Installing Updates:"
foreach ($update in $updates) {
  Write-Host $update.Title
}

$result = $installer.Install()  

Write-Host "Successfully installed updates:"
foreach ($installed in $result.Updates) {
  if($installed.ResultCode -eq 0) {    
    Write-Host $installed.Title 
  }
}

$optionals = $searcher.Search("IsInstalled=0 and IsHidden=0 and IsAssigned=0").Updates 
if($optionals.Count -gt 0) {
  Write-Host "Optional updates not selected for install:"
  foreach ($optional in $optionals) { 
     Write-Host $optional.Title
  } 
}
