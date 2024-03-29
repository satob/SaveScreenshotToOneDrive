﻿# Get from https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps
$ClientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$RedirectURI = "http://localhost/SaveScreenShot"
# Should be "https://xxxxx-my.sharepoint.com/"
$ResourceId = "https://xxxxx-my.sharepoint.com/"
# You have to save when you register the app
$AppKey = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"

# Give "Files.ReadWrite.All" and "Files.Read.All" permission to this application 


[void] [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [Reflection.Assembly]::LoadWithPartialName("System.Web")

# Tasktray Icon
$notify_icon = $null

function Invoke-WebRequestThruProxy {
  PARAM(
    [Parameter(Mandatory = $True)]  $Method,
    [Parameter(Mandatory = $True)]  [uri]$Uri,
    [Parameter(Mandatory = $False)] $ContentType = "application/x-www-form-urlencoded",
    [Parameter(Mandatory = $False)] $Body = $null,
    [Parameter(Mandatory = $False)] $InFile = $null,
    [Parameter(Mandatory = $False)] $Header = $null,
    [Parameter(Mandatory = $False)] $TimeoutSec = 10
  )

  $WithProxy = $false
  $BaseUri = $Uri.Scheme + "://" + $Uri.Host
  $ProxyUri = [System.Net.WebRequest]::GetSystemWebProxy().GetProxy($BaseUri)
  if (([uri]$BaseUri).AbsoluteUri -ne $ProxyUri.AbsoluteUri) {
    $WithProxy = $true
  }

  try {
    if ($null -ne $Body) {
      if ($WithProxy) {
        $Response = Invoke-WebRequest -Method $Method -Uri $Uri -ContentType $ContentType -Body $Body -Header $Header -UseBasicParsing -Proxy $ProxyUri.AbsoluteUri -ProxyUseDefaultCredentials -TimeoutSec $TimeoutSec
      }
      else {
        $Response = Invoke-WebRequest -Method $Method -Uri $Uri -ContentType $ContentType -Body $Body -Header $Header -UseBasicParsing -TimeoutSec $TimeoutSec
      }
    }
    elseif ($null -ne $InFile) {
      if ($WithProxy) {
        $Response = Invoke-WebRequest -Method $Method -Uri $Uri -ContentType $ContentType -InFile $InFile -Header $Header -UseBasicParsing -Proxy $ProxyUri.AbsoluteUri -ProxyUseDefaultCredentials -TimeoutSec $TimeoutSec
      }
      else {
        $Response = Invoke-WebRequest -Method $Method -Uri $Uri -ContentType $ContentType -InFile $InFile -Header $Header -UseBasicParsing -TimeoutSec $TimeoutSec
      }
    }
    return ($Response.Content | ConvertFrom-Json)
  }
  catch {
    $e = $_.Exception
    $Line = $_.InvocationInfo.PositionMessage
    $Message = $e.Message
    $Parameters = ($PSBoundParameters.Keys | ForEach-Object { $_ + "=" + $PSBoundParameters.Item($_) }) -join "`n"
    # [void][System.Windows.Forms.MessageBox]::Show("At $Line" + "`n" + $Message + "`n" + $Parameters, "Error", "OK", "Information")
    $notify_icon.BalloonTipIcon  = [Windows.Forms.ToolTipIcon]::Warning
    # $notify_icon.BalloonTipTitle = 'Error'
    $notify_icon.BalloonTipText = "At $Line" + "`n" + $Message + "`n" + $Parameters
    $notify_icon.ShowBalloonTip(60000)
    $error.clear()
  }
}


function Authenticate-OneDrive {
  PARAM(
    [Parameter(Mandatory = $True)]
    [string]$ClientId,
    [string]$Scope = "offline_access%20User.Read%20Files.ReadWrite.All%20Files.Read.All",
    [string]$RedirectURI = "https://login.live.com/oauth20_desktop.srf",
    [string]$AppKey = "",
    [string]$ResourceId = "",
    [string]$State
  )

  $AuthorizeURI = "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize?client_id=${ClientId}&response_type=code&redirect_uri=${RedirectURI}&response_mode=query&scope=${Scope}&state=${State}"

  $Form = New-Object Windows.Forms.Form
  $Form.text = "Authenticate to OneDrive"
  $Form.size = New-Object Drawing.size @(840, 525)
  $Form.Height = 840
  $Form.Width = 525
  $Web = New-object System.Windows.Forms.WebBrowser
  $Web.IsWebBrowserContextMenuEnabled = $true
  $Web.Height = 800
  $Web.Width = 500
  $Web.Location = "0, 0"
  $Web.navigate($AuthorizeURI)

  $DocComplete = {
    if ($Web.Url.AbsoluteUri -match "access_token=|error|code=|logout") {
      $Form.Close()
    }
    if ($Web.DocumentText -like '*ucaccept*') {
      $Web.Document.GetElementById("idBtn_Accept").InvokeMember("click")
    }
  }
  $Web.Add_DocumentCompleted($DocComplete)
  $Form.Controls.Add($Web)
  $Form.showdialog() | Out-Null

  $ReturnURI = ($Web.Url).ToString().Replace("#", "&")

  $Authentication = New-Object PSObject
  ForEach ($Element in $ReturnURI.Split("?")[1].Split("&")) {
    $Authentication | Add-Member Noteproperty $Element.split("=")[0] $Element.split("=")[1]
  }
  if ($Authentication.code) {
    $Code = $Authentication.code
    $Body = "client_id=${ClientId}&redirect_URI=${RedirectURI}&code=${Code}&grant_type=authorization_code&scope=${Scope}"
    return Invoke-WebRequestThruProxy -Method POST -Uri "https://login.microsoftonline.com/organizations/oauth2/v2.0/token" -Body $Body
  }
  else {
    Write-Error ("Cannot get authentication code. Error: " + $ReturnURI) -ErrorAction Stop
  }
}


function Refresh-Token {
  PARAM(
    [Parameter(Mandatory = $True)]
    [string]$ClientId,
    [string]$RedirectURI = "https://login.live.com/oauth20_desktop.srf",
    [string]$AppKey = "",
    [string]$RefreshToken = ""
  )
  $Body = "client_id=${ClientId}&redirect_URI=${RedirectURI}&refresh_token=${RefreshToken}&grant_type=refresh_token"
  return Invoke-WebRequestThruProxy -Method POST -Uri "https://login.microsoftonline.com/organizations/oauth2/v2.0/token" -Body $Body
}


function Upload-OneDrive {
  PARAM(
    [Parameter(Mandatory = $True)]
    [string]$AccessToken,
    [string]$Path = "/",
    [string]$LocalFile = "",
    [string]$ContentType = "application/octet-stream"
  )
  $RootURI = "https://graph.microsoft.com/v1.0/me/drive/root"
  $Header = @{ Authorization = "Bearer " + $AccessToken; Host = "graph.microsoft.com" }
  return Invoke-WebRequestThruProxy -Method PUT -Uri ($RootURI + ":" + $Path + ":/content") -InFile $LocalFile -Header $Header -ContentType $ContentType
}


# Framework from https://qiita.com/magiclib/items/cc2de9169c781642e52d
Add-Type -AssemblyName System.Windows.Forms

# 定数定義
$TIMER_INTERVAL = 300 * 1000 # timer_function実行間隔(ミリ秒)
$MUTEX_NAME = "Global\mutex" # 多重起動チェック用

function timer_function($notify) {
  # トークンのリフレッシュ
  $Authentication_Old = $Authentication
  $Authentication = Refresh-Token -ClientId $ClientId -RedirectURI $RedirectURI -AppKey $AppKey -RefreshToken $Authentication.refresh_token
  if ($null -eq $Authentication) {
    $Authentication = $Authentication_Old
  }

  try {
    $Year = "{0:0000}" -F (Get-Date).Year
    $Month = "{0:00}" -F (Get-Date).Month
    $Day = "{0:00}" -F (Get-Date).Day
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

    foreach ($item in $screen) {
      if ($workingAreaX -gt $item.WorkingArea.X) {
        $workingAreaX = $item.WorkingArea.X;
      }
      if ($workingAreaY -gt $item.WorkingArea.Y) {
        $workingAreaY = $item.WorkingArea.Y;
      }
      $width = $width + $item.Bounds.Width;

      if ($item.Bounds.Height -gt $height) {
        $height = $item.Bounds.Height;
      }
    }

    $bounds = [Drawing.Rectangle]::FromLTRB($workingAreaX, $workingAreaY, $width, $height); 
    $bmp = New-Object Drawing.Bitmap $width, $height;
    $graphics = [Drawing.Graphics]::FromImage($bmp);

    # It fails when the screen is locked
    $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size);

    # 文字用ブラシ（アルファ値100）
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(200, 128, 128, 128))
    # フォントの指定
    $font = New-Object System.Drawing.Font("Arial Black", 36) 
    # 左50px、上50pxに描画
    $graphics.DrawString((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $font, $brush, 16, 16)

    $bmp.Save($TmpFilePath, $JpegEncoder, $JpegEncoderParameters);
    Upload-OneDrive -AccessToken $Authentication.access_token -Path "/Screenshot/${Year}/${Month}/${Day}/${FileName}" -LocalFile $TmpFilePath -ContentType "image/jpeg"
  }
  catch {
    # DO NOTHING
  }

  $graphics.Dispose();
  $bmp.Dispose();

  $datetime = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
  Write-Host "timer_function "  $datetime
}


function Hide-ConsoleWindow() {
  $ShowWindowAsyncCode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
  $ShowWindowAsync = Add-Type -MemberDefinition $ShowWindowAsyncCode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru

  $hwnd = (Get-Process -PID $pid).MainWindowHandle
  if ($hwnd -ne [System.IntPtr]::Zero) {
    # When you got HWND of the console window:
    # (It would appear that Windows Console Host is the default terminal application)
    $ShowWindowAsync::ShowWindowAsync($hwnd, 0)
  } else {
    # When you failed to get HWND of the console window:
    # (It would appear that Windows Terminal is the default terminal application)

    # Mark the current console window with a unique string.
    $UniqueWindowTitle = New-Guid
    $Host.UI.RawUI.WindowTitle = $UniqueWindowTitle
    $StringBuilder = New-Object System.Text.StringBuilder 1024

    # Search the process that has the window title generated above.
    $TerminalProcess = (Get-Process | Where-Object { $_.MainWindowTitle -eq $UniqueWindowTitle })
    # Get the window handle of the terminal process.
    # Note that GetConsoleWindow() in Win32 API returns the HWND of
    # powershell.exe itself rather than the terminal process.
    # When you call ShowWindowAsync(HWND, 0) with the HWND from GetConsoleWindow(),
    # the Windows Terminal window will be just minimized rather than hidden.
    $hwnd = $TerminalProcess.MainWindowHandle
    if ($hwnd -ne [System.IntPtr]::Zero) {
      $ShowWindowAsync::ShowWindowAsync($hwnd, 0)
    } else {
      Write-Host "Failed to hide the console window."
    }
  }
}


function main() {
  $mutex = New-Object System.Threading.Mutex($false, $MUTEX_NAME)
  # 多重起動チェック
  if ($mutex.WaitOne(0, $false)) {
    Hide-ConsoleWindow

    $application_context = New-Object System.Windows.Forms.ApplicationContext
    $timer = New-Object Windows.Forms.Timer

    # タスクトレイアイコンの画像を作成
    $FormIconGZipBase64 = 'H4sIAAAAAAAEAMVUP2gTURz+xKHFf83kFEgHEUGH4pRBSBAEQcQigYzJ5CbJ5OBywcE4BK4iuL4MVW+7Tma8GwKigxdwUVHubRkyXDar1X5+l7Q0jWkiFPRdvrvX+97vd9/ve79X4ISu1VXovgxvGTgP4JKgVyhi/H7+aIx+aDTGj/TSJAxDhF9CDN+91NtTwmnhjHBWOCesTHylKjhCWwgFK+wimyXyeaJUIup1otUiPI/odglruSgcWWaRZx4lllBnHS224NFDl11Y2kXhYDYL5vNgqQTW62CrBXoe2O2CVqtS44rKUFUGRxna7bRwJRC3K/2UfqEk1IWW4AldwQoLwp+Q2etkPqT0U/op/ZR+Sj9H+lelvyj9Vel3pL8t/aH0W+nflX5KvyD/KP8o/yj/KP8o/+64ESomRs1P4ASE+5YwHwj/KxH05W+xgVzZRaFmUGn6cEwA04kQRDHifgJqn+nmQFMA/QoYOGBkwDgAkxgN1eaWczC1AvxmBYFxEHUM4ihA0o9B9QpzLlhQTMUHHcWZSHnExYnYItxcGaZQg19pInAMItNBHERI4j5SATlKH6WP0kfpo/RR+pggFZBzyyiYGip+UzUamKiDII4QJ31cvkdce0DcfkxUnmuPXhGPOsSzN8SLj1Kn+l3Vb1S/r/oD1R+p/lj1J6q/IYtcqj6qPqo+qj5piaUjkQaFoyx/avKnKX+M/OnIn0j+9OUP7qrJbl0AbtwECveB/FPg6mvgymfgopKvCw+FTSLznihuEzX1TVvoCZlMRj20irW1NfVREevr6+qlqs5KXf3kYGNjQz3VxtbW1ug89no99ZbFcDhUf6Udvi38En4KO8IP4TuGO8Sm4hrpQdYupT813PiRXumr9PYvxmDmmM8erFjAT/xtR2NqwQR7+DnF270Ee9Np3u4Hpnc7i7eDP6YH/ETQxHySn9Bt/wc/X9+i+hb6M+XveHL0/szgD+3vAX/c/jlywaH2xQxqMJcfpMTsQ7DHHnVK9pNPLzi8eh7594MTQ/9JchZYShECJ4+D/TxD5Rw2Tn4TPjFcqtGudPSpYP+bvwGD5A+AvggAAA=='
    $FormIconGZip = [System.Convert]::FromBase64String($FormIconGZipBase64)
    $FormIconGZipMemoryStream = New-Object System.IO.MemoryStream(, $FormIconGZip)
    $FormIconMemoryStream = New-Object System.IO.MemoryStream
    $GZipStream = New-Object System.IO.Compression.GzipStream $FormIconGZipMemoryStream, ([IO.Compression.CompressionMode]::Decompress)
    $GZipStream.CopyTo( $FormIconMemoryStream )
    $GZipStream.Close()
    $FormIconGZipMemoryStream.Close()
    $icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($FormIconMemoryStream).GetHIcon()))

    # タスクトレイアイコンを設定
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


function Test-NetConnectionToMS365 {
  $Uri = [uri]"https://login.microsoftonline.com/common/v2.0/.well-known/openid-configuration"

  $WithProxy = $false
  $BaseUri = $Uri.Scheme + "://" + $Uri.Host
  $ProxyUri = [System.Net.WebRequest]::GetSystemWebProxy().GetProxy($BaseUri)
  if (([uri]$BaseUri).AbsoluteUri -ne $ProxyUri.AbsoluteUri) {
    $WithProxy = $true
  }

  try {
    if ($WithProxy) {
      $Response = Invoke-WebRequest -Method GET -Uri $Uri -UseBasicParsing -Proxy $ProxyUri.AbsoluteUri -ProxyUseDefaultCredentials -TimeoutSec 3
    }
    else {
      $Response = Invoke-WebRequest -Method GET -Uri $Uri -UseBasicParsing -TimeoutSec 3
    }
    ($Response.Content | ConvertFrom-Json)
    return $true
  } catch {
    return $false
  }
}

# ネットワークが有効になるまで待つ
while ($true) {
  $PingHosts | ForEach-Object {
    if (Test-NetConnectionToMS365) {
      break;
    }
    Start-Sleep -Seconds 10
  }
}

# まず認証する
try {
  $State = (Get-Random)
  $Authentication = Authenticate-OneDrive -ClientId $ClientId -AppKey $AppKey -RedirectURI $RedirectURI -ResourceId $ResourceId
}
catch {
  [void][System.Windows.Forms.MessageBox]::Show($error, "Error", "OK", "Information")
  $error
}

main
