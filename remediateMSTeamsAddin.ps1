# $currentMSTeamsPaths = @("C:\Program Files\WindowsApps\MSTeams_23335.232.2637.4844_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi", "C:\Program Files\WindowsApps\MSTeams_24004.1305.2651.7623_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi")
# $regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect"

$ErrorActionPreference = "SilentlyContinue"
$regTeamsAddin = "HKCU:\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect\"

$obj = Get-WmiObject -Class win32_product -Filter "Name = 'Microsoft Teams Meeting Add-in for Microsoft Office'"

if($null -ne $obj){
    write-output "Uninstalling $($obj.Name)"
    $obj.Uninstall() | Out-Null
}

$proc = Get-Process -Name ms-teams, OUTLOOK
$sProc = stop-process -InputObject $proc -Force

$tp = Start-Process -FilePath ms-teams.exe -PassThru

do{
    # write-output "waiting for the teams process to start."
    Start-Sleep -Seconds 3
    if($null -ne $tp){ break }

}while($null -eq $tp)

Start-Sleep -Seconds 5

if((Get-ItemPropertyValue -Path $regTeamsAddin -Name "LoadBehavior") -eq 3){
    write-output "$($obj.Name) installed correctly. Exit 0"; exit 0
}


Write-Output "Teams Meeting Addin failed to load. exit 1"; Exit 1


# $msiDetected = $false 

# $currentMSTeamsPaths | ForEach-Object {
#     if($_ -contains ""){}
# }

# if($msiDetected -eq $false){
#     Write-Output "MicrosoftTeamsMeetingAddinInstaller.msi does not exist. Exit 0"; exit 0
# }
    
# if( -not (Test-Path "$($env:LOCALAPPDATA)\Microsoft\TeamsMeetingAddin")){
#     Write-Output "Teams Addin path not detected. Exit 1"; exit 1
# }

# if( -not (test-path $regPath) ){
#     Write-Output "$regPath not detected. Exit 1"; exit 1
# }

# $regValue = Get-ItemPropertyValue -path "HKCU:\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" -Name "LoadBehavior" -ErrorAction SilentlyContinue

# if($regValue -ne 3){
#     Write-Output "$regPath contains WRONG value of $regValue. Exit 1"; exit 1
# }

# Write-Output "TeamsAddin.FastConnect is installed properly. Exit 0"; exit 0