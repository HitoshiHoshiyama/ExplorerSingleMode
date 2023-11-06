Param([Parameter(Mandatory=$false)][string]$Exedir = $PSScriptRoot)

$ExePath = Convert-Path (Join-Path $Exedir "ExplorerSingleMode.exe")
try {
    $TaskAction = New-ScheduledTaskAction -Id "ExplorerSingleMode" -Execute $ExePath
}
catch {
    Write-Host エラー発生のため、タスクを登録しませんでした。
    exit
}
$TaskTrigger = New-ScheduledTaskTrigger -AtLogOn
$TaskOptions = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Seconds 0) -Hidden

Write-Host タスク:ExplorerSingleMode の実行ファイル[$ExePath]

$IsExist = Get-ScheduledTask | Select-String ExplorerSingleMode
if(-not [string]::IsNullOrEmpty($IsExist))
{
    . .\TaskRemove.ps1 -NoWait
}

Register-ScheduledTask -TaskName "ExplorerSingleMode" -Action $TaskAction -Trigger $TaskTrigger -Settings $TaskOptions -RunLevel "Highest" -Description "エクスプローラのタブをひとつのウィンドウにまとめます。"
Write-Host タスク:ExplorerSingleMode を登録しました。

Start-ScheduledTask -TaskName "ExplorerSingleMode"
Write-Host タスク:ExplorerSingleMode を開始しました。

Get-ScheduledTask -TaskName "ExplorerSingleMode"

$KeyIn = Read-Host "何かキーを押すと終了します。"
Write-Host $KeyIn