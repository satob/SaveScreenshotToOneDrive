# Get from https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps
$ClientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$RedirectURI = "http://localhost/SaveScreenShot"
# Should be "https://xxxxx-my.sharepoint.com/"
$ResourceId = "https://xxxxx-my.sharepoint.com/"
# You have to save when you register the app
$AppKey = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"

# Give "Files.ReadWrite.All" permission to this application 


# functions from https://github.com/MarcelMeurer/PowerShellGallery-OneDrive
function Get-ODAuthentication
{
	<#
	.DESCRIPTION
	Connect to OneDrive for authentication with a given client id (get your free client id on https://apps.dev.microsoft.com) For a step-by-step guide: https://github.com/MarcelMeurer/PowerShellGallery-OneDrive
	.PARAMETER ClientId
	ClientId of your "app" from https://apps.dev.microsoft.com
	.PARAMETER AppKey
	The client secret for your OneDrive "app". If AppKey is set the authentication mode is "code." Code authentication returns a refresh token to refresh your authentication token unattended.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER Scope
	Comma-separated string defining the authentication scope (https://dev.onedrive.com/auth/msa_oauth.htm). Default: "onedrive.readwrite,offline_access". Not needed for OneDrive 4 Business access.
	.PARAMETER RefreshToken
	Refreshes the authentication token unattended with this refresh token. 
	.PARAMETER AutoAccept
	In token mode the accept button in the web form is pressed automatically.
	.PARAMETER RedirectURI
	Code authentication requires a correct URI. Use the same as in the app registration e.g. http://localhost/logon. Default is https://login.live.com/oauth20_desktop.srf. Don't use this parameter for token-based authentication. 

	.EXAMPLE
    $Authentication=Get-ODAuthentication -ClientId "0000000012345678"
	$AuthToken=$Authentication.access_token
	Connect to OneDrive for authentication and save the token to $AuthToken
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$True)]
		[string]$ClientId = "unknown",
		[string]$Scope = "onedrive.readwrite,offline_access",
		[string]$RedirectURI ="https://login.live.com/oauth20_desktop.srf",
		[string]$AppKey="",
		[string]$RefreshToken="",
		[string]$ResourceId="",
		[switch]$DontShowLoginScreen=$false,
		[switch]$AutoAccept,
		[switch]$LogOut
	)
	$optResourceId=""
	$optOauthVersion="/v2.0"
	if ($ResourceId -ne "")
	{
		write-debug("Running in OneDrive 4 Business mode")
		$optResourceId="&resource=$ResourceId"
		$optOauthVersion=""
	}
	$Authentication=""
	if ($AppKey -eq "")
	{ 
		$Type="token"
	} else 
	{ 
		$Type="code"
	}
	
	if ($RefreshToken -ne "")
	{
		write-debug("A refresh token is given. Try to refresh it in code mode.")
		$body="client_id=$ClientId&redirect_URI=$RedirectURI&client_secret=$([uri]::EscapeDataString($AppKey))&refresh_token="+$RefreshToken+"&grant_type=refresh_token"
		if ($ResourceId -ne "")
		{
			# OD4B
			$webRequest=Invoke-WebRequest -Method POST -Uri "https://login.microsoftonline.com/common/oauth2$optOauthVersion/token" -ContentType "application/x-www-form-urlencoded" -Body $Body -UseBasicParsing
		} else {
			# OD private
			$webRequest=Invoke-WebRequest -Method POST -Uri "https://login.live.com/oauth20_token.srf" -ContentType "application/x-www-form-urlencoded" -Body $Body -UseBasicParsing
		}
		$Authentication = $webRequest.Content |   ConvertFrom-Json
	} else
	{
		write-debug("Authentication mode: " +$Type)
		[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | out-null
		[Reflection.Assembly]::LoadWithPartialName("System.Drawing") | out-null
		[Reflection.Assembly]::LoadWithPartialName("System.Web") | out-null
		if ($Logout)
		{
			$URIGetAccessToken="https://login.live.com/logout.srf"
		}
		else
		{
			if ($ResourceId -ne "")
			{
				# OD4B
				$URIGetAccessToken="https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&client_id=$ClientId&redirect_URI=$RedirectURI"
			}
			else
			{
				# OD private
				$URIGetAccessToken="https://login.live.com/oauth20_authorize.srf?client_id="+$ClientId+"&scope="+$Scope+"&response_type="+$Type+"&redirect_URI="+$RedirectURI
			}
		}
		$form = New-Object Windows.Forms.Form
		$form.text = "Authenticate to OneDrive"
		$form.size = New-Object Drawing.size @(700,600)
		$form.Width = 675
		$form.Height = 750
		$web=New-object System.Windows.Forms.WebBrowser
		$web.IsWebBrowserContextMenuEnabled = $true
		$web.Width = 600
		$web.Height = 700
		$web.Location = "25, 25"
		$web.navigate($URIGetAccessToken)
		$DocComplete  = {
			if ($web.Url.AbsoluteUri -match "access_token=|error|code=|logout") {$form.Close() }
			if ($web.DocumentText -like '*ucaccept*') {
				if ($AutoAccept) {$web.Document.GetElementById("idBtn_Accept").InvokeMember("click")}
			}
		}
		$web.Add_DocumentCompleted($DocComplete)
		$form.Controls.Add($web)
		if ($DontShowLoginScreen)
		{
			write-debug("Logon screen suppressed by flag -DontShowLoginScreen")
			$form.Opacity = 0.0;
		}
		$form.showdialog() | out-null
		# Build object from last URI (which should contains the token)
		$ReturnURI=($web.Url).ToString().Replace("#","&")
		if ($LogOut) {return "Logout"}
		if ($Type -eq "code")
		{
			write-debug("Getting code to redeem token")
			$Authentication = New-Object PSObject
			ForEach ($element in $ReturnURI.Split("?")[1].Split("&")) 
			{
				$Authentication | add-member Noteproperty $element.split("=")[0] $element.split("=")[1]
			}
			if ($Authentication.code)
			{
				$body="client_id=$ClientId&redirect_URI=$RedirectURI&client_secret=$([uri]::EscapeDataString($AppKey))&code="+$Authentication.code+"&grant_type=authorization_code"+$optResourceId+"&scope="+$Scope
			if ($ResourceId -ne "")
			{
				# OD4B
				$webRequest=Invoke-WebRequest -Method POST -Uri "https://login.microsoftonline.com/common/oauth2$optOauthVersion/token" -ContentType "application/x-www-form-urlencoded" -Body $Body -UseBasicParsing
			} else {
				# OD private
				$webRequest=Invoke-WebRequest -Method POST -Uri "https://login.live.com/oauth20_token.srf" -ContentType "application/x-www-form-urlencoded" -Body $Body -UseBasicParsing
			}
			$Authentication = $webRequest.Content |   ConvertFrom-Json
			} else
			{
				write-error("Cannot get authentication code. Error: "+$ReturnURI)
			}
		} else
		{
			$Authentication = New-Object PSObject
			ForEach ($element in $ReturnURI.Split("?")[1].Split("&")) 
			{
				$Authentication | add-member Noteproperty $element.split("=")[0] $element.split("=")[1]
			}
			if ($Authentication.PSobject.Properties.name -match "expires_in")
			{
				$Authentication | add-member Noteproperty "expires" ([System.DateTime]::Now.AddSeconds($Authentication.expires_in))
			}
		}
	}
	if (!($Authentication.PSobject.Properties.name -match "expires_in"))
	{
		write-warning("There is maybe an errror, because there is no access_token!")
	}
	return $Authentication 
}

function Get-ODRootUri 
{
	PARAM(
		[String]$ResourceId=""
	)
	if ($ResourceId -ne "")
	{
		return $ResourceId+"_api/v2.0"
	}
	else
	{
		return "https://api.onedrive.com/v1.0"
	}
}

function Get-ODWebContent 
{
	<#
	.DESCRIPTION
	Internal function to interact with the OneDrive API
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER rURI
	Relative path to the API.
	.PARAMETER Method
	Webrequest method like PUT, GET, ...
	.PARAMETER Body
	Payload of a webrequest.
	.PARAMETER BinaryMode
	Do not convert response to JSON.
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$True)]
		[string]$AccessToken,
		[String]$ResourceId="",
		[string]$rURI = "",
		[ValidateSet("PUT","GET","POST","PATCH","DELETE")] 
        [String]$Method="GET",
		[String]$Body,
		[switch]$BinaryMode
	)
	if ($Body -eq "") 
	{
		$xBody=$null
	} else
	{
		$xBody=$Body
	}
	$ODRootURI=Get-ODRootUri -ResourceId $ResourceId
	try {
		$webRequest=Invoke-WebRequest -Method $Method -Uri ($ODRootURI+$rURI) -Header @{ Authorization = "BEARER "+$AccessToken} -ContentType "application/json" -Body $xBody -UseBasicParsing -ErrorAction SilentlyContinue
	} 
	catch
	{
		write-error("Cannot access the api. Webrequest return code is: "+$_.Exception.Response.StatusCode+"`n"+$_.Exception.Response.StatusDescription)
		break
	}
	switch ($webRequest.StatusCode) 
    { 
        200 
		{
			if (!$BinaryMode) {$responseObject = ConvertFrom-Json $webRequest.Content}
			return $responseObject
		} 
        201 
		{
			write-debug("Success: "+$webRequest.StatusCode+" - "+$webRequest.StatusDescription)
			if (!$BinaryMode) {$responseObject = ConvertFrom-Json $webRequest.Content}
			return $responseObject
		} 
        204 
		{
			write-debug("Success: "+$webRequest.StatusCode+" - "+$webRequest.StatusDescription+" (item deleted)")
			$responseObject = "0"
			return $responseObject
		} 
        default {write-warning("Cannot access the api. Webrequest return code is: "+$webRequest.StatusCode+"`n"+$webRequest.StatusDescription)}
    }
}

function Get-ODDrives
{
	<#
	.DESCRIPTION
	Get user's drives.
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.EXAMPLE
    Get-ODDrives -AccessToken $AuthToken
	List all OneDrives available for your account (there is normally only one).
	.NOTES
	The application for OneDrive 4 Business needs "Read items in all site collections" on application level (API: Office 365 SharePoint Online)
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$True)]
		[string]$AccessToken,
		[String]$ResourceId=""
	)
	$ResponseObject=Get-ODWebContent -AccessToken $AccessToken -ResourceId $ResourceId -Method GET -rURI "/drives" 
	return $ResponseObject.Value
}

function Get-ODSharedItems
{
	<#
	.DESCRIPTION
	Get items shared with the user
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.EXAMPLE
    Get-ODDrives -AccessToken $AuthToken
	List all OneDrives available for your account (there is normally only one).
	.NOTES
	The application for OneDrive 4 Business needs "Read items in all site collections" on application level (API: Office 365 SharePoint Online)
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$True)]
		[string]$AccessToken,
		[String]$ResourceId=""
	)
	$ResponseObject=Get-ODWebContent -AccessToken $AccessToken -ResourceId $ResourceId -Method GET -rURI "/drive/oneDrive.sharedWithMe"
	return $ResponseObject.Value
}

function Format-ODPathorIdStringV2
{
	<#
	.DESCRIPTION
	Formats a given path like '/myFolder/mySubfolder/myFile' into an expected URI format
	.PARAMETER Path
	Specifies the path of an element. If it is not given, the path is "/"
	.PARAMETER ElementId
	Specifies the id of an element. If Path and ElementId are given, the ElementId is used with a warning
	.PARAMETER DriveId
	Specifies the OneDrive drive id. If not set, the default drive is used
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[string]$Path="",
		[string]$DriveId="",
		[string]$ElementId=""
	)
	if (!$ElementId -eq "")
	{
		# Use ElementId parameters
		if (!$Path -eq "") {write-debug("Warning: Path and ElementId parameters are set. Only ElementId is used!")}
		$drive="/drive"
		if ($DriveId -ne "") 
		{	
			# Named drive
			$drive="/drives/"+$DriveId
		}
		return $drive+"/items/"+$ElementId
	}
	else
	{
		# Use Path parameter
		# replace some special characters
		$Path = ((((($Path -replace '%', '%25') -replace ' ', ' ') -replace '=', '%3d') -replace '\+', '%2b') -replace '&', '%26') -replace '#', '%23'
		# remove substring starts with "?"
		if ($Path.Contains("?")) {$Path=$Path.Substring(1,$Path.indexof("?")-1)}
		# replace "\" with "/"
		$Path=$Path.Replace("\","/")
		# filter possible string at the end "/children" (case insensitive)
		$Path=$Path+"/"
		$Path=$Path -replace "/children/",""
		# encoding of URL parts
		$tmpString=""
		foreach ($Sub in $Path.Split("/")) {$tmpString+=$Sub+"/"}
		$Path=$tmpString
		# remove last "/" if exist 
		$Path=$Path.TrimEnd("/")
		# insert drive part of URL
		if ($DriveId -eq "") 
		{	
			# Default drive
			$Path="/drive/root:"+$Path+""
		}
		else
		{
			# Named drive
			$Path="/drives/"+$DriveId+"/root:"+$Path+":"
		}
		return ($Path).replace("root::","root:")
	}
}

function Format-ODPathorIdString
{
	<#
	.DESCRIPTION
	Formats a given path like '/myFolder/mySubfolder/myFile' into an expected URI format
	.PARAMETER Path
	Specifies the path of an element. If it is not given, the path is "/"
	.PARAMETER ElementId
	Specifies the id of an element. If Path and ElementId are given, the ElementId is used with a warning
	.PARAMETER DriveId
	Specifies the OneDrive drive id. If not set, the default drive is used
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[string]$Path="",
		[string]$DriveId="",
		[string]$ElementId=""
	)
	if (!$ElementId -eq "")
	{
		# Use ElementId parameters
		if (!$Path -eq "") {write-debug("Warning: Path and ElementId parameters are set. Only ElementId is used!")}
		$drive="/drive"
		if ($DriveId -ne "") 
		{	
			# Named drive
			$drive="/drives/"+$DriveId
		}
		return $drive+"/items/"+$ElementId
	}
	else
	{
		# Use Path parameter
		# replace some special characters
		$Path = ((((($Path -replace '%', '%25') -replace ' ', ' ') -replace '=', '%3d') -replace '\+', '%2b') -replace '&', '%26') -replace '#', '%23'
		# remove substring starts with "?"
		if ($Path.Contains("?")) {$Path=$Path.Substring(1,$Path.indexof("?")-1)}
		# replace "\" with "/"
		$Path=$Path.Replace("\","/")
		# filter possible string at the end "/children" (case insensitive)
		$Path=$Path+"/"
		$Path=$Path -replace "/children/",""
		# encoding of URL parts
		$tmpString=""
		foreach ($Sub in $Path.Split("/")) {$tmpString+=$Sub+"/"}
		$Path=$tmpString
		# remove last "/" if exist 
		$Path=$Path.TrimEnd("/")
		# insert drive part of URL
		if ($DriveId -eq "") 
		{	
			# Default drive
			$Path="/drive/root:"+$Path+":"
		}
		else
		{
			# Named drive
			$Path="/drives/"+$DriveId+"/root:"+$Path+":"
		}
		return ($Path).replace("root::","root")
	}
}

function Get-ODItemProperty
{
	<#
	.DESCRIPTION
	Get the properties of an item (file or folder).
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER Path
	Specifies the path to the element/item. If not given, the properties of your default root drive are listed.
	.PARAMETER ElementId
	Specifies the id of the element/item. If Path and ElementId are given, the ElementId is used with a warning.
	.PARAMETER SelectProperties
	Specifies a comma-separated list of the properties to be returned for file and folder objects (case sensitive). If not set, name, size, lastModifiedDateTime and id are used. (See https://dev.onedrive.com/odata/optional-query-parameters.htm).
	If you use -SelectProperties "", all properties are listed. Warning: A complex "content.downloadUrl" is listed/generated for download files without authentication for several hours.
	.PARAMETER DriveId
	Specifies the OneDrive drive id. If not set, the default drive is used.
	.EXAMPLE
    Get-ODItemProperty -AccessToken $AuthToken -Path "/Data/documents/2016/AzureML with PowerShell.docx"
	Get the default set of metadata for a file or folder (name, size, lastModifiedDateTime, id)

	Get-ODItemProperty -AccessToken $AuthToken -ElementId 8BADCFF017EAA324!12169 -SelectProperties ""
	Get all metadata of a file or folder by element id ("" select all properties)	
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$True)]
		[string]$AccessToken,
		[string]$ResourceId="",
		[string]$Path="/",
		[string]$ElementId="",
		[string]$SelectProperties="name,size,lastModifiedDateTime,id",
		[string]$DriveId=""
	)
	return Get-ODChildItems -AccessToken $AccessToken -ResourceId $ResourceId -Path $Path -ElementId $ElementId -SelectProperties $SelectProperties -DriveId $DriveId -ItemPropertyMode
}

function Get-ODChildItems
{
	<#
	.DESCRIPTION
	Get child items of a path. Return count is not limited.
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER Path
	Specifies the path of elements to be listed. If not given, the path is "/".
	.PARAMETER ElementId
	Specifies the id of an element. If Path and ElementId are given, the ElementId is used with a warning.
	.PARAMETER SelectProperties
	Specifies a comma-separated list of the properties to be returned for file and folder objects (case sensitive). If not set, name, size, lastModifiedDateTime and id are used. (See https://dev.onedrive.com/odata/optional-query-parameters.htm).
	If you use -SelectProperties "", all properties are listed. Warning: A complex "content.downloadUrl" is listed/generated for download files without authentication for several hours.
	.PARAMETER DriveId
	Specifies the OneDrive drive id. If not set, the default drive is used.
	.EXAMPLE
    Get-ODChildItems -AccessToken $AuthToken -Path "/" | ft
	Lists files and folders in your OneDrives root folder and displays name, size, lastModifiedDateTime, id and folder property as a table

    Get-ODChildItems -AccessToken $AuthToken -Path "/" -SelectProperties ""
	Lists files and folders in your OneDrives root folder and displays all properties
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$True)]
		[string]$AccessToken,
		[String]$ResourceId="",
		[string]$Path="/",
		[string]$ElementId="",
		[string]$SelectProperties="name,size,lastModifiedDateTime,id",
		[string]$DriveId="",
		[Parameter(DontShow)]
		[switch]$ItemPropertyMode,
		[Parameter(DontShow)]
		[string]$SearchText,
		[parameter(DontShow)]
        [switch]$Loop=$false
	)
	$ODRootURI=Get-ODRootUri -ResourceId $ResourceId
	if ($Path.Contains('$skiptoken=') -or $Loop)
	{	
		# Recursive mode of odata.nextLink detection
		write-debug("Recursive call")
		$rURI=$Path	
	}
	else
	{
		$rURI=Format-ODPathorIdString -path $Path -ElementId $ElementId -DriveId $DriveId
		$rURI=$rURI.Replace("::","")
		$SelectProperties=$SelectProperties.Replace(" ","")
		if ($SelectProperties -eq "")
		{
			$opt=""
		} else
		{
			$SelectProperties=$SelectProperties.Replace(" ","")+",folder"
			$opt="?select="+$SelectProperties
		}
		if ($ItemPropertyMode)
		{
			# item property mode
			$rURI=$rURI+$opt
		}
		else
		{
			if (!$SearchText -eq "") 
			{
				# Search mode
				$opt="/view.search?q="+$SearchText+"&select="+$SelectProperties
				$rURI=$rURI+$opt
			}
			else
			{
				# child item mode
				$rURI=$rURI+"/children"+$opt
			}
		}
	}
	write-debug("Accessing API with GET to "+$rURI)
	$ResponseObject=Get-ODWebContent -AccessToken $AccessToken -ResourceId $ResourceId -Method GET -rURI $rURI
	if ($ResponseObject.PSobject.Properties.name -match "@odata.nextLink") 
	{
		write-debug("Getting more elements form service (@odata.nextLink is present)")
		write-debug("LAST: "+$ResponseObject.value.count)
		Get-ODChildItems -AccessToken $AccessToken -ResourceId $ResourceId -SelectProperties $SelectProperties -Path $ResponseObject."@odata.nextLink".Replace($ODRootURI,"") -Loop
	}
	if ($ItemPropertyMode)
	{
		# item property mode
		return $ResponseObject
	}
	else
	{
		# child item mode
		return $ResponseObject.value
	}
}

function New-ODFolder
{
	<#
	.DESCRIPTION
	Create a new folder.
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER FolderName
	Name of the new folder.
	.PARAMETER Path
	Specifies the parent path for the new folder. If not given, the path is "/".
	.PARAMETER ElementId
	Specifies the element id for the new folder. If Path and ElementId are given, the ElementId is used with a warning.
	.PARAMETER DriveId
	Specifies the OneDrive drive id. If not set, the default drive is used.
	.EXAMPLE
    New-ODFolder -AccessToken $AuthToken -Path "/data/documents" -FolderName "2016"
	Creates a new folder "2016" under "/data/documents"
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$True)]
		[string]$AccessToken,
		[String]$ResourceId="",
		[Parameter(Mandatory=$True)]
		[string]$FolderName,
		[string]$Path="/",
		[string]$ElementId="",
		[string]$DriveId=""
	)
	$rURI=Format-ODPathorIdString -path $Path -ElementId $ElementId -DriveId $DriveId
	$rURI=$rURI+"/children"
	return Get-ODWebContent -AccessToken $AccessToken -ResourceId $ResourceId -Method POST -rURI $rURI -Body ('{"name": "'+$FolderName+'","folder": { },"@name.conflictBehavior": "fail"}')
}

function Add-ODItem
{
	<#
	.DESCRIPTION
	Upload an item/file. Warning: An existing file will be overwritten.
	.PARAMETER AccessToken
	A valid access token for bearer authorization.
	.PARAMETER ResourceId
	Mandatory for OneDrive 4 Business access. Is the ressource URI: "https://<tenant>-my.sharepoint.com/". Example: "https://sepagogmbh-my.sharepoint.com/"
	.PARAMETER Path
	Specifies the path for the upload folder. If not given, the path is "/".
	.PARAMETER ElementId
	Specifies the element id for the upload folder. If Path and ElementId are given, the ElementId is used with a warning.
	.PARAMETER DriveId
	Specifies the OneDrive drive id. If not set, the default drive is used.
	.PARAMETER LocalFile
	Path and file of the local file to be uploaded (C:\data\data.csv).
	.EXAMPLE
    Add-ODItem -AccessToken $AuthToken -Path "/Data/documents/2016" -LocalFile "AzureML with PowerShell.docx" 
    Upload a file to OneDrive "/data/documents/2016"
	.NOTES
    Author: Marcel Meurer, marcel.meurer@sepago.de, Twitter: MarcelMeurer
	#>
	PARAM(
		[Parameter(Mandatory=$True)]
		[string]$AccessToken,
		[String]$ResourceId="",
		[string]$Path="/",
		[string]$ElementId="",
		[string]$DriveId="",
		[Parameter(Mandatory=$True)]
		[string]$LocalFile=""
	)
	$rURI=Format-ODPathorIdString -path $Path -ElementId $ElementId -DriveId $DriveId
	try
	{
		$spacer=""
		if ($ElementId -ne "") {$spacer=":"}
		$ODRootURI=Get-ODRootUri -ResourceId $ResourceId
		$rURI=(($ODRootURI+$rURI).TrimEnd(":")+$spacer+"/"+[System.IO.Path]::GetFileName($LocalFile)+":/content").Replace("/root/","/root:/")
		return $webRequest=Invoke-WebRequest -Method PUT -InFile $LocalFile -Uri $rURI -Header @{ Authorization = "BEARER "+$AccessToken} -ContentType "multipart/form-data"  -UseBasicParsing -ErrorAction SilentlyContinue
	}
	catch
	{
		write-error("Upload error: "+$_.Exception.Response.StatusCode+"`n"+$_.Exception.Response.StatusDescription)
		return -1
	}	
}



# Framework from https://qiita.com/magiclib/items/cc2de9169c781642e52d
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

  try {
    # It fails when the screen is locked
    $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size);

    # 文字用ブラシ（アルファ値100）
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(200, 128, 128, 128))
    # フォントの指定
    $font = New-Object System.Drawing.Font("Arial Black", 36) 
    # 左50px、上50pxに描画
    $graphics.DrawString((Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $font, $brush, 16, 16)

    $bmp.Save($TmpFilePath, $JpegEncoder, $JpegEncoderParameters);
    Add-ODItem -AccessToken $Authentication.access_token -ResourceId $ResourceId -LocalFile $TmpFilePath -Path "/Screenshot/${Year}/${Month}/${Day}"
  } catch {
    # DO NOTHING
  }

  $graphics.Dispose();
  $bmp.Dispose();

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


# ネットワークが有効になるまで待つ
$PingHosts = @('login.live.com', 'login.microsoftonline.com');
while ($true) {
  $PingHosts | ForEach-Object {
    $AuthFQDN = $_;
    $AuthIP = (Resolve-DnsName $AuthFQDN | Where-Object { $_.QueryType -eq "A" } | Select-Object -First 1).IPAddress;
    if ((Test-NetConnection -ComputerName $AuthIP -Port 443).TcpTestSucceeded) {
      break;
    }
    Start-Sleep -Seconds 10
  }
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
