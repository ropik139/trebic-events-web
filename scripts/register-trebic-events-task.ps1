[CmdletBinding()]
param(
    [string]$TaskName = "DailyTrebicEventsMonitor",
    [string]$RunTime = "15:00"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptPath = [System.IO.Path]::GetFullPath((Join-Path (Split-Path -Parent $PSCommandPath) "update-trebic-events.ps1"))

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
$trigger = New-ScheduledTaskTrigger -Daily -At $RunTime
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Denní přehled kulturních akcí pro Třebíč a okolí do 10 km"

Write-Host "Naplánovaná úloha '$TaskName' je vytvořena a bude se spouštět denně v $RunTime."
