# Get from https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps
$ClientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$RedirectURI = "http://localhost/SaveScreenShot"
# Should be "https://xxxxx-my.sharepoint.com/"
$ResourceId = "https://xxxxx-my.sharepoint.com/"
# You have to save when you register the app
$AppKey = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"

# まず認証する
$Authentication = Get-ODAuthentication -ClientId $ClientId -AppKey $AppKey -RedirectURI $RedirectURI -ResourceId $ResourceId

# OneDriveにスクリーンショット用フォルダがなければ作る
$ErrorActionPreference = "Stop"
try {
  Get-ODChildItems -AccessToken $Authentication.access_token -ResourceId $ResourceId -Path "/Screenshot"
} catch {
  New-ODFolder -AccessToken $Authentication.access_token -ResourceId $ResourceId -Path "/" -FolderName "Screenshot"
}

# 本日のスクリーンショット用フォルダがなければ作る
$Year  = "{0:0000}" -F (Get-Date).Year
$Month = "{0:00}" -F (Get-Date).Month
$Day   = "{0:00}" -F (Get-Date).Day

try {
  Get-ODChildItems -AccessToken $Authentication.access_token -ResourceId $ResourceId -Path "/Screenshot/${Year}"
} catch {
  New-ODFolder -AccessToken $Authentication.access_token -ResourceId $ResourceId -Path "/Screenshot" -FolderName $Year
}

try {
  Get-ODChildItems -AccessToken $Authentication.access_token -ResourceId $ResourceId -Path "/Screenshot/${Year}/${Month}"
} catch {
  New-ODFolder -AccessToken $Authentication.access_token -ResourceId $ResourceId -Path "/Screenshot/${Year}" -FolderName $Month
}

try {
  Get-ODChildItems -AccessToken $Authentication.access_token -ResourceId $ResourceId -Path "/Screenshot/${Year}/${Month}/${Day}"
} catch {
  New-ODFolder -AccessToken $Authentication.access_token -ResourceId $ResourceId -Path "/Screenshot/${Year}/${Month}" -FolderName $Day
}

while ($true) {
  # トークンのリフレッシュ
  $LastAuthentication = $Authentication
  $Authentication = Get-ODAuthentication -ClientId $ClientId -AppKey $AppKey -RedirectURI $RedirectURI -ResourceId $ResourceId -RefreshToken $LastAuthentication.refresh_token

  # スクリーンショットを取る
  Add-Type -AssemblyName System.Windows.Forms

  $FileName = (Get-Date).ToString('HHmm') + ".jpg"
  $TmpFilePath = Join-Path $env:TEMP $FileName
  $JpegEncoder = ([System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.FormatID -eq [System.Drawing.Imaging.ImageFormat]::Jpeg.Guid })
  $JpegEncoderParameters = New-Object System.Drawing.Imaging.EncoderParameters(1)
  $JpegEncoderParameters.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, 10)

  $b = New-Object System.Drawing.Bitmap([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width, [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height)
  $g = [System.Drawing.Graphics]::FromImage($b)
  $g.CopyFromScreen((New-Object System.Drawing.Point(0,0)), (New-Object System.Drawing.Point(0,0)), $b.Size)
  $g.Dispose()
  $b.Save($TmpFilePath, $JpegEncoder, $JpegEncoderParameters)

  Add-ODItem -AccessToken $Authentication.access_token -ResourceId $ResourceId -LocalFile $TmpFilePath -Path "/Screenshot/${Year}/${Month}/${Day}"

  Start-Sleep -Seconds 300
}

