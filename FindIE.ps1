Param(
    [Int]$hour,
    [String]$path,
    [Int]$nolog
)

if(-Not(Test-Path $Path)){
    New-Item $Path -ItemType Directory
}

$StartDate = (Get-Date) - (New-TimeSpan -Hours $hour)
$Event = Get-WinEvent -FilterHashtable @{
    Logname = 'Security'
    ID = 4688
    StartTime = $StartDate
}

$Count = 0

ForEach ($i in $Event){
    $EventXML = [XML]($i.ToXML())
    $ParentProc = ($EventXML.Event.EventData.Data | where Name -eq "ParentProcessName")."#text"
    $ParentProc = Split-Path $ParentProc -Leaf
    $NewProc = ($EventXML.Event.EventData.Data | where Name -eq "NewProcessName")."#text"
    $NewProc = Split-Path $NewProc -Leaf
    
    if(($ParentProc -eq "explorer.exe") -and ($NewProc -eq "iexplore.exe")){
        $output += "ParentProc:" + $ParentProc + "`r`n"
        $output += "NewProc:" + $NewProc + "`r`n"
        $output += ($i.TimeCreated.ToString("yyyy/MM/dd HH:mm:ss")) + "`r`n"
        $output += $EventXML.Event.System.Computer + "`r`n"
        $output += ($EventXML.Event.EventData.Data | where Name -eq "SubjectUserName")."#text" + "`r`n" + "`r`n"
        $Count += 1
    }
}
if(-Not($Count -eq 0)){
    $output += $StartDate.ToString("yyyy/MM/dd HH:mm:ss") + "から過去" + $hour + "時間の間に" + $Count + "件のIE起動イベントログがありました。"
    $FileName = $StartDate.ToString("yyyyMMdd-HHmmss") + "_" + $env:COMPUTERNAME + ".txt"
    $output | Out-File -FilePath ($path + "`\" + $FileName)
}elseif($nolog -eq 1){
    $output += "IE実行のイベントログは見つかりませんでした。"
    $FileName = $StartDate.ToString("yyyyMMdd-HHmmss") + "_" + $env:COMPUTERNAME + "_NotFound" + ".txt"
    $output | Out-File -FilePath ($path + "`\" + $FileName)
}