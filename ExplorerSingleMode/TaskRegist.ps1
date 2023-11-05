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
    write-Host タスク:ExplorerSingleMode が既に存在するので削除します。
    Unregister-ScheduledTask -TaskName "ExplorerSingleMode" -Confirm:$false
}

Register-ScheduledTask -TaskName "ExplorerSingleMode" -Action $TaskAction -Trigger $TaskTrigger -Settings $TaskOptions
Write-Host タスク:ExplorerSingleMode を登録しました。

$KeyIn = Read-Host "何かキーを押すと終了します。"
Write-Host $KeyIn