Param([Parameter(Mandatory=$false)][switch]$NoWait)

$IsExist = Get-ScheduledTask | Select-String ExplorerSingleMode
if(-not [string]::IsNullOrEmpty($IsExist))
{
    $Task = Get-ScheduledTask -TaskName ExplorerSingleMode
    if(($Task).State -eq 'Running')
    {
        write-Host タスク:ExplorerSingleMode を停止します。
        Stop-ScheduledTask -TaskName "ExplorerSingleMode"
    }
    write-Host タスク:ExplorerSingleMode を削除します。
    Unregister-ScheduledTask -TaskName "ExplorerSingleMode" -Confirm:$false
}else {
    write-Host タスク:ExplorerSingleMode は登録されていません。
}

if(-not $NoWait)
{
    $KeyIn = Read-Host "何かキーを押すと終了します。"
    Write-Host $KeyIn
}