Add-Type -AssemblyName System.Drawing
$img = [System.Drawing.Image]::FromFile('D:\Personal Praison\Develop\Asset manager\AssetVault_app_icon.png')
$bmp = New-Object System.Drawing.Bitmap(64,64)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.DrawImage($img, 0, 0, 64, 64)
$ms = New-Object System.IO.MemoryStream
$bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
[Convert]::ToBase64String($ms.ToArray()) | Out-File -FilePath "$env:TEMP\small_icon.txt"
$ico = [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
$fs = New-Object System.IO.FileStream('D:\Personal Praison\Develop\Asset manager\AssetVault.ico', [System.IO.FileMode]::Create)
$ico.Save($fs)
$fs.Close()
