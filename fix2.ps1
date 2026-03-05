$Content = Get-Content 'D:\Personal Praison\Develop\Asset manager\AssetManagerPro.ps1' -Raw

$Delimiter1 = '        $PdfFile = "$DownloadsPath\HardwareReport_${ComputerName}_$Timestamp.pdf"'
$idx1 = $Content.IndexOf($Delimiter1)
if ($idx1 -ge 0) {
    $TopBlock = $Content.Substring(0, $idx1 + $Delimiter1.Length)
}

$Delimiter2 = '        $HtmlBody | Out-File -FilePath $HtmlFile -Encoding UTF8'
$idx2 = $Content.LastIndexOf($Delimiter2)
if ($idx2 -ge 0) {
    $BotBlock = $Content.Substring($idx2)
}

$InjectContent = Get-Content 'D:\Personal Praison\Develop\Asset manager\inject_code.ps1' -Raw
$InjectContent -match '(?ms)\$ReplaceBlock = @''\r?\n(.*?)''@\r?\n'
$MiddleBlock = $matches[1]

if ($TopBlock -and $BotBlock -and $MiddleBlock) {
    $NewContent = $TopBlock + "`r`n`r`n" + $MiddleBlock + "`r`n`r`n" + $BotBlock
    $NewContent | Out-File -FilePath 'D:\Personal Praison\Develop\Asset manager\AssetManagerPro.ps1' -Encoding UTF8
    Write-Host "Fixed asset manager script."
} else {
    Write-Host "Failed to find delimiters."
    Write-Host "Top: $([bool]$TopBlock), Bot: $([bool]$BotBlock), Mid: $([bool]$MiddleBlock)"
}
