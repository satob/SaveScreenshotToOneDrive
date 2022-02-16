# Get from https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps
$ClientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$RedirectURI = "http://localhost/SaveScreenShot"
# Should be "https://xxxxx-my.sharepoint.com/"
$ResourceId = "https://xxxxx-my.sharepoint.com/"
# You have to save when you register the app
$AppKey = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"

# Give "Files.ReadWrite.All" permission to this application 


Add-Type -AssemblyName System.Windows.Forms

# 定数定義
$TIMER_INTERVAL = 300 * 1000 # timer_function実行間隔(ミリ秒)
$MUTEX_NAME = "Global\mutex" # 多重起動チェック用

function timer_function($notify){
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


  $datetime = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
  Write-Host "timer_function "  $datetime
}

function main(){
  $mutex = New-Object System.Threading.Mutex($false, $MUTEX_NAME)
  # 多重起動チェック
  if ($mutex.WaitOne(0, $false)){
    # タスクバー非表示
    $windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
    $asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
    $null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

    $application_context = New-Object System.Windows.Forms.ApplicationContext
    $timer = New-Object Windows.Forms.Timer
    $path = Get-Process -id $pid | Select-Object -ExpandProperty Path # icon用
    $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)

    $FormIconGZipBase64 = 'H4sIAAAAAAAEAMVUP2gTURz+xKHFf83kFEgHEUGH4pRBSBAEQcQigYzJ5CbJ5OBywcE4BK4iuL4MVW+7Tma8GwKigxdwUVHubRkyXDar1X5+l7Q0jWkiFPRdvrvX+97vd9/ve79X4ISu1VXovgxvGTgP4JKgVyhi/H7+aIx+aDTGj/TSJAxDhF9CDN+91NtTwmnhjHBWOCesTHylKjhCWwgFK+wimyXyeaJUIup1otUiPI/odglruSgcWWaRZx4lllBnHS224NFDl11Y2kXhYDYL5vNgqQTW62CrBXoe2O2CVqtS44rKUFUGRxna7bRwJRC3K/2UfqEk1IWW4AldwQoLwp+Q2etkPqT0U/op/ZR+Sj9H+lelvyj9Vel3pL8t/aH0W+nflX5KvyD/KP8o/yj/KP8o/+64ESomRs1P4ASE+5YwHwj/KxH05W+xgVzZRaFmUGn6cEwA04kQRDHifgJqn+nmQFMA/QoYOGBkwDgAkxgN1eaWczC1AvxmBYFxEHUM4ihA0o9B9QpzLlhQTMUHHcWZSHnExYnYItxcGaZQg19pInAMItNBHERI4j5SATlKH6WP0kfpo/RR+pggFZBzyyiYGip+UzUamKiDII4QJ31cvkdce0DcfkxUnmuPXhGPOsSzN8SLj1Kn+l3Vb1S/r/oD1R+p/lj1J6q/IYtcqj6qPqo+qj5piaUjkQaFoyx/avKnKX+M/OnIn0j+9OUP7qrJbl0AbtwECveB/FPg6mvgymfgopKvCw+FTSLznihuEzX1TVvoCZlMRj20irW1NfVREevr6+qlqs5KXf3kYGNjQz3VxtbW1ug89no99ZbFcDhUf6Udvi38En4KO8IP4TuGO8Sm4hrpQdYupT813PiRXumr9PYvxmDmmM8erFjAT/xtR2NqwQR7+DnF270Ee9Np3u4Hpnc7i7eDP6YH/ETQxHySn9Bt/wc/X9+i+hb6M+XveHL0/szgD+3vAX/c/jlywaH2xQxqMJcfpMTsQ7DHHnVK9pNPLzi8eh7594MTQ/9JchZYShECJ4+D/TxD5Rw2Tn4TPjFcqtGudPSpYP+bvwGD5A+AvggAAA=='
    $FormIconGZip = [System.Convert]::FromBase64String($FormIconGZipBase64)
    $FormIconGZipMemoryStream = New-Object System.IO.MemoryStream(, $FormIconGZip)
    $FormIconMemoryStream = New-Object System.IO.MemoryStream
    $GZipStream = New-Object System.IO.Compression.GzipStream $FormIconGZipMemoryStream, ([IO.Compression.CompressionMode]::Decompress)
    $GZipStream.CopyTo( $FormIconMemoryStream )
    $GZipStream.Close()
    $FormIconGZipMemoryStream.Close()
    $icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($FormIconMemoryStream).GetHIcon()))

    # タスクトレイアイコン
    $notify_icon = New-Object System.Windows.Forms.NotifyIcon
    $notify_icon.Icon = $icon
    $notify_icon.Visible = $true
    $notify_icon.Text = "スクリーンショット保存君"

    # アイコンクリック時のイベント
    $notify_icon.add_Click({
      if ($_.Button -eq [Windows.Forms.MouseButtons]::Left) {
        # タイマーで実装されているイベントを即時実行する
        $timer.Stop()
        $timer.Interval = 1
        $timer.Start()
      }
    })

    # メニュー
    $menu_item_exit = New-Object System.Windows.Forms.MenuItem
    $menu_item_exit.Text = "Exit"
    $notify_icon.ContextMenu = New-Object System.Windows.Forms.ContextMenu
    $notify_icon.contextMenu.MenuItems.AddRange($menu_item_exit)

    # Exitメニュークリック時のイベント
    $menu_item_exit.add_Click({
      $application_context.ExitThread()
    })

    # タイマーイベント.
    $timer.Enabled = $true
    $timer.Add_Tick({
      $timer.Stop()

      timer_function($notify_icon)

      # インターバルを再設定してタイマー再開
      $timer.Interval = $TIMER_INTERVAL
      $timer.Start()
    })

    $timer.Interval = 1
    $timer.Start()

    [void][System.Windows.Forms.Application]::Run($application_context)

    $timer.Stop()
    $notify_icon.Visible = $false
    $mutex.ReleaseMutex()
  }
  $mutex.Close()
}


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

# スクリーンショットを取る準備
[Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

main
