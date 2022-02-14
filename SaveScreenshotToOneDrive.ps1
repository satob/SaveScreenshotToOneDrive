# Get from https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps
$ClientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$RedirectURI = "http://localhost/SaveScreenShot"
# Should be "https://xxxxx-my.sharepoint.com/"
$ResourceId = "https://xxxxx-my.sharepoint.com/"
# You have to save when you register the app
$AppKey = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"

# Give "Files.ReadWrite.All" permission to this application 


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


# コマンドプロンプトを隠す
Add-Type -Name Window -Namespace Console -MemberDefinition @'
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'@
Start-Sleep -Seconds 1
$ConsolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($ConsolePtr, 0)
Start-Sleep -Seconds 20


# スクリーンショットを取る
[Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

while ($true) {
  # トークンのリフレッシュ
  $LastAuthentication = $Authentication
  $Authentication = Get-ODAuthentication -ClientId $ClientId -AppKey $AppKey -RedirectURI $RedirectURI -ResourceId $ResourceId -RefreshToken $LastAuthentication.refresh_token


  $FileName = (Get-Date).ToString('HHmm') + ".jpg"
  $TmpFilePath = Join-Path $env:TEMP $FileName
  $JpegEncoder = ([System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.FormatID -eq [System.Drawing.Imaging.ImageFormat]::Jpeg.Guid })
  $JpegEncoderParameters = New-Object System.Drawing.Imaging.EncoderParameters(1)
  $JpegEncoderParameters.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, 10)

  $width = 0;
  $height = 0;
  $workingAreaX = 0;
  $workingAreaY = 0;

  $screen = [System.Windows.Forms.Screen]::AllScreens;

  foreach ($item in $screen)
  {
    if($workingAreaX -gt $item.WorkingArea.X) {
      $workingAreaX = $item.WorkingArea.X;
    }
    if($workingAreaY -gt $item.WorkingArea.Y) {
      $workingAreaY = $item.WorkingArea.Y;
    }
    $width = $width + $item.Bounds.Width;

    if($item.Bounds.Height -gt $height) {
      $height = $item.Bounds.Height;
    }
  }

  $bounds = [Drawing.Rectangle]::FromLTRB($workingAreaX, $workingAreaY, $width, $height); 
  $bmp = New-Object Drawing.Bitmap $width, $height;
  $graphics = [Drawing.Graphics]::FromImage($bmp);

  $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size);

  $bmp.Save($TmpFilePath, $JpegEncoder, $JpegEncoderParameters);

  $graphics.Dispose();
  $bmp.Dispose();

  Add-ODItem -AccessToken $Authentication.access_token -ResourceId $ResourceId -LocalFile $TmpFilePath -Path "/Screenshot/${Year}/${Month}/${Day}"

  Start-Sleep -Seconds 300
}

