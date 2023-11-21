<#
Author: Damien VAN ROBAEYS
Website: https://www.systanddeploy.com
Twitter: @syst_and_deploy
Mail: damien.vanrobaeys@gmail.com
#>

<#
Prerequisites for the purge
1. Create a SharePoint application
2. fill SharePoint app in below variables
Check my post here below:
https://www.systanddeploy.com/2022/02/how-to-use-teamssharepoint-as-logs.html

Prerequisites for the Teams notification
1. Create a webhook on a Teams channel (see below)
2. Add the webhook URL in variable Webhook_URL

To create a webhook proceed as below:
1. Go to your channel
2. Click on the ...
3. Click on Connectors
4. Go to Incoming Webhook
5. Type a name
6. Click on Create
7. Copy the Webhook path
#>

<#
Getting Sharepoint site id
I have the following Sharepoint site: https://systanddeploy.sharepoint.com/sites/Support
In order to authenticate and upload file we need to get the id of the site.
For this just open your browser and type:
https://m365x53191121.sharepoint.com/sites/systanddeploy/_api/site/id
#>

# Information about Teams webhook
$Webhook_URL = ""	

# info abut SharePoint
$SharePoint_Path = ""  # sharepoint main path
$SharePoint_ExportFolder = ""  # folder where to upload file	

<#
Example
$SharePoint_Path = "https://systanddeploy.sharepoint.com/sites/Support"  # sharepoint main path
$SharePoint_ExportFolder = "Windows/Logs"  # folder where to upload file
#>
	
# Function used to send notif on Teams
Function Send_Notif
	{
			param(
			$Text,	
			$Title
			)

			$Body = @{
			'text'= $Text
			'Title'= $Title
			'themeColor'= $MessageColor
			}

			$Params = @{
					 Headers = @{'accept'='application/json'}
					 Body = $Body | ConvertTo-Json
					 Method = 'Post'
					 URI = $Webhook_URL 
			}
			Invoke-RestMethod @Params
	}	

Function Write_Log
	{
		param(
		$Message_Type,	
		$Message
		)
		
		$MyDate = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)		
		# Add-Content $Log_File  "$MyDate - $Message_Type : $Message"	
		write-host  "$MyDate - $Message_Type : $Message"			
	}	

$Current_Folder = split-path $MyInvocation.MyCommand.Path						

Try
	{
		$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
		$Script:ClientID = $tsenv.Value("TS_Client_ID") 	
		$Script:Secret = $tsenv.Value("TS_Client_Secret") 	
		$Script:SharePoint_SiteID = $tsenv.Value("TS_SharePoint_Site_ID") 	
		$Script:Tenant = $tsenv.Value("TS_Tenant_Name") 	
		$Script:Send_Teams_Notif = $tsenv.Value("TS_Send_Teams_Notif") 			
		$Script:LastActionName = $tsenv.Value("LastActionName") 		
	}
Catch
	{
		EXIT
	}

$Log_to_send_ZIP = "C:\Windows\Temp\Logs_$env:computername.zip"
$Logs_Export_folder = "C:\Windows\Temp\Logs_$env:computername"
If(test-path $Logs_Export_folder){remove-item $Logs_Export_folder -Force -Recurse}
new-item $Logs_Export_folder -Type Directory -Force | out-null
If(test-path $Log_to_send_ZIP){remove-item $Log_to_send_ZIP -Force}

Function Copy_Logs{
	param(
	$Logs_Path_To_Copy,
	$SubFolder
	)
	If(Test-Path -Path $Logs_Path_To_Copy)
		{
			If($SubFolder -ne $null)
				{
					$Logs_Copy_Path = "$Logs_Export_folder\$SubFolder"
					new-item $Logs_Copy_Path -Type Directory -Force | out-null
				}
			Else
				{
					$Logs_Copy_Path = "$Logs_Export_folder"
				}
			copy-item $Logs_Path_To_Copy $Logs_Copy_Path -recurse -force
		}
	}		
		
#*************************************************************
# 						Capturing logs
#*************************************************************	
	
# System logs
$DISM_folder = "$env:SystemRoot\Logs\DISM"
Copy_Logs -Logs_Path_To_Copy $DISM_folder -SubFolder "DISM"

$Panther_folder = "C:\Windows\Panther"
Copy_Logs -Logs_Path_To_Copy $Panther_folder -SubFolder "Panther"

$Win_Debug_folder = "$env:SystemRoot\debug"
Copy_Logs -Logs_Path_To_Copy $Win_Debug_folder -SubFolder "Debug"

# $logfiles_folder = "C:\Windows\System32\logfiles"
# Copy_Logs -Logs_Path_To_Copy $logfiles_folder #-SubFolder "logfiles"

# Event logs
$EVTX_Logs_Folder = "$Logs_Export_folder\Event logs"
If(!(test-path $EVTX_Logs_Folder)){new-item $EVTX_Logs_Folder -type directory -force | out-null}
wevtutil epl system "$EVTX_Logs_Folder\System_logs.evtx"
wevtutil epl Application "$EVTX_Logs_Folder\Application_logs.evtx"
wevtutil epl Setup "$EVTX_Logs_Folder\Setup_logs.evtx" 	

# Deployment Logs
Try
	{
		$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
		$Logs_Folder = $tsenv.Value("LogPath") 	
		Copy_Logs -Logs_Path_To_Copy $Logs_Folder -SubFolder "TS_Logs"
	}
Catch{}

$CCM_logs = "$env:SystemRoot\CCM\Logs"
# Copy_Logs -Logs_Path_To_Copy $SMSOSD_Temp_folder -SubFolder "CCM_Logs"

$DeploymentLogs_folder = "$env:SystemRoot\Temp\DeploymentLogs"
Copy_Logs -Logs_Path_To_Copy $DeploymentLogs_folder -SubFolder "DeploymentLogs"

$MININT_folder = "$env:SystemDrive\MININT"
Copy_Logs -Logs_Path_To_Copy $MININT_folder -SubFolder "MININT"

$SMSTaskSequence_folder = "$env:SystemDrive\_SMSTaskSequence\Logs"
Copy_Logs -Logs_Path_To_Copy $SMSTaskSequence_folder -SubFolder "_SMSTaskSequence"

$SMSTaskSequence_C_folder = "C:\_SMSTaskSequence\Logs"
Copy_Logs -Logs_Path_To_Copy $SMSTaskSequence_C_folder -SubFolder "_SMSTaskSequence_on_C"

$SMSTSLogs_folder = "x:\smstslog"
Copy_Logs -Logs_Path_To_Copy $SMSTSLogs_folder -SubFolder "smstslog"

$SMSTSLogs_temp_folder = "x:\temp\smstslog"
Copy_Logs -Logs_Path_To_Copy $SMSTSLogs_temp_folder -SubFolder "smstslog_temp"

$SMSTS_PowerShellScripts_temp_folder = "x:\temp\smstsPowerShellScripts"
Copy_Logs -Logs_Path_To_Copy $SMSTS_PowerShellScripts_temp_folder -SubFolder "smstsPowerShellScripts"

$SMSOSD_folder = "$env:SystemRoot\SMSOSD"
Copy_Logs -Logs_Path_To_Copy $SMSOSD_folder -SubFolder "SMSOSD"

$SMSOSD_Temp_folder = "$env:SystemRoot\TEMP\SMSOSD"
Copy_Logs -Logs_Path_To_Copy $SMSOSD_Temp_folder -SubFolder "Temp_SMSOSD"


#*************************************************************
# 					Creating a ZIP file with logs
#*************************************************************	
Compress-Archive -Path $Logs_Export_folder -DestinationPath $Log_to_send_ZIP -Update


Log_to_send_ZIP






# Authentification sur SharePoint et upload du fichier
$Body = @{  
    client_id = $ClientID
    client_secret = $Secret
    scope = "https://graph.microsoft.com/.default"   
    grant_type = 'client_credentials'  
}  
  
Write_Log -Message_Type "INFO" -Message "SharePoint connexion"	
$Graph_Url = "https://login.microsoftonline.com/$($Tenant).onmicrosoft.com/oauth2/v2.0/token"  

Try
	{
		$AuthorizationRequest = Invoke-RestMethod -Uri $Graph_Url -Method "Post" -Body $Body  
		Write_Log -Message_Type "SUCCESS" -Message "Connected to SharePoint"	
	}
Catch
	{
		Write_Log -Message_Type "ERROR" -Message "Connexion to SharePoint failed"	
		EXIT
	}
	
$Access_token = $AuthorizationRequest.Access_token  
$Header = @{  
    Authorization = $AuthorizationRequest.access_token  
    "Content-Type"= "application/json"  
    'Content-Range' = "bytes 0-$($fileLength-1)/$fileLength"	
}  

$SharePoint_Graph_URL = "https://graph.microsoft.com/v1.0/sites/$SharePoint_SiteID/drives"  
$BodyJSON = $Body | ConvertTo-Json -Compress  

Write_Log -Message_Type "INFO" -Message "Getting SharePoint site info"	

Try
	{
		$Result = Invoke-RestMethod -Uri $SharePoint_Graph_URL -Method 'GET' -Headers $Header -ContentType "application/json"   
		Write_Log -Message_Type "SUCCESS" -Message "Getting SharePoint site info"		
	}
Catch
	{
		Write_Log -Message_Type "ERROR" -Message "Getting SharePoint site info"	
		EXIT
	}

$DriveID = $Result.value| Where-Object {$_.webURL -eq $SharePoint_Path } | Select-Object id -ExpandProperty id  

$FileName = $Log_to_send_ZIP.Split("\")[-1]  
$createUploadSessionUri = "https://graph.microsoft.com/v1.0/sites/$SharePoint_SiteID/drives/$DriveID/root:/$SharePoint_ExportFolder/$($fileName):/createUploadSession"

Write_Log -Message_Type "INFO" -Message "File to upload: $FileName"	
Write_Log -Message_Type "INFO" -Message "Preparing the file for the upload"	

Try
	{
		$uploadSession = Invoke-RestMethod -Uri $createUploadSessionUri -Method 'POST' -Headers $Header -ContentType "application/json" 
		Write_Log -Message_Type "SUCCESS" -Message "Preparing the file for the upload"			
	}
Catch
	{
		Write_Log -Message_Type "ERROR" -Message "Preparing the file for the upload"			
		EXIT
	}

$fileInBytes = [System.IO.File]::ReadAllBytes($Log_to_send_ZIP)
$fileLength = $fileInBytes.Length

$headers = @{
  'Content-Range' = "bytes 0-$($fileLength-1)/$fileLength"
}

$Upload_Status = $false
Write_Log -Message_Type "INFO" -Message "Uploading file"	
Try
	{
		$response = Invoke-RestMethod -Method 'Put' -Uri $uploadSession.uploadUrl -Body $fileInBytes -Headers $headers
		Write_Log -Message_Type "SUCCESS" -Message "File has been uploaded"	
		$Upload_Status = $true
	}
Catch
	{
		Write_Log -Message_Type "ERROR" -Message "Failed to upload the file"
		EXIT
	}

	
If(($Send_Teams_Notif -eq $True) -and ($Upload_Status -eq $True))
	{
		$Title_Message = "Task Sequence has failed on device $env:computername"
		$Text_Message = "<b>ZIP name</b>: Logs_$env:computername.zip<br>
		"
		
		Send_Notif -Text $Text_Message -Title $Title_Message | out-null			
	}	