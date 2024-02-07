
# $currentMSTeamsPaths = @("C:\Program Files\WindowsApps\MSTeams_23335.232.2637.4844_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi", "C:\Program Files\WindowsApps\MSTeams_24004.1305.2651.7623_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi")

$regPath = "HKCU:\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect"
$teamsAddinLocalPath = "$($env:LOCALAPPDATA)\Microsoft\TeamsMeetingAddin"

if($null -eq (Get-AppxPackage -name MSTeams)){
    Write-Output "New MSTeams AppxPackage was not detected. Exit 0"; exit 0
}
    
if( -not (Test-Path $teamsAddinLocalPath) ){
    Write-Output "Teams Addin path not detected. Exit 1"; exit 1
}

if( -not (test-path $regPath) ){
    Write-Output "$regPath not detected. Exit 1"; exit 1
}

$regValue = Get-ItemPropertyValue -path "HKCU:\Software\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect" -Name "LoadBehavior" -ErrorAction SilentlyContinue

if($regValue -ne 3){
    Write-Output "$regPath contains WRONG value of $regValue. Exit 1"; exit 1
}

Write-Output "TeamsAddin.FastConnect is installed properly. Exit 0"; exit 0







