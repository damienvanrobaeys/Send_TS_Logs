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


# Information about Teams webhook
$Webhook_URL = ""

# Info about SharePoint application
$Sharepoint_Folder = ""
$Sharepoint_Site_URL = ""		
	
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
$PnP_Module_Path = "$Current_Folder\PnP.PowerShell"

Try
	{
		$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
		$Script:Sharepoint_App_ID = $tsenv.Value("TS_Sharepoint_App_ID") 	
		$Script:Sharepoint_App_Secret = $tsenv.Value("TS_Sharepoint_App_Secret") 				
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

#*************************************************************
# 					Installing module
#*************************************************************	
$Is_Nuget_Installed = $False     
$PnP_Module_Status = $False

If(!(Get-PackageProvider | where {$_.Name -eq "Nuget"}))
	{                                         
		Try
			{
				[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
				Install-PackageProvider -Name Nuget -MinimumVersion 2.8.5.201 -Scope currentuser -Force -Confirm:$False | out-null                                                                                                                 
				$Is_Nuget_Installed = $True 
				Write_Log -Message_Type "INFO" -Message "Package Nuget installé"				
			}
		Catch
			{
				$Is_Nuget_Installed = $False  
				Write_Log -Message_Type "INFO" -Message "Package Nuget installé"								
			}
	}
Else
	{
		$Is_Nuget_Installed = $True      
	}

# If($Is_Nuget_Installed -eq $True)
	# {	
		# Try
			# {
				# Import-Module $PnP_Module_Path -Force -ErrorAction SilentlyContinue 
				# $PnP_Module_Status = $True	  
				# Write_Log -Message_Type "SUCCESS" -Message "Module PnP importé"						
			# }
		# Catch
			# {
				# Write_Log -Message_Type "ERROR" -Message "Module PnP importé"										
			# }                                                     
	# }

#*************************************************************
# 						Sending logs
#*************************************************************
# If($PnP_Module_Status -eq $True)
	# { 
		Try
			{
				Connect-PnPOnline -Url $Sharepoint_Site_URL -ClientID $Sharepoint_App_ID -ClientSecret $Sharepoint_App_Secret -WarningAction Ignore		
				$Sharepoint_Status = "OK"
				Write_Log -Message_Type "SUCCESS" -Message "Connexion SharePoint"								
			}
		Catch
			{
				$Sharepoint_Status = "KO"	
				Write_Log -Message_Type "ERROR" -Message "Connexion SharePoint"												
			}	 
	
		If($Sharepoint_Status -eq "OK")
			{
				Write_Log -Message_Type "INFO" -Message "Upload du fichier"								

				Try
					{
						Write_Log -Message_Type "INFO" -Message "Upload du fichier en cours"													
						Add-PnPFile -Path $Log_to_send_ZIP -Folder $Sharepoint_Folder | out-null					
						Write_Log -Message_Type "SUCCESS" -Message "Upload du fichier"		
						$Upload_Status = $True						
					}
				Catch
					{
						Write_Log -Message_Type "ERROR" -Message "Upload du fichier"
						$Last_Error = $error[0]
						Write_Log -Message_Type "ERROR" -Message "$Last_Error"						
						$Upload_Status = $False												
					}
			}	
	# }
	
	
	
If(($Send_Teams_Notif -eq $True) -and ($Upload_Status -eq $True))
	{
		$Title_Message = "Task Sequence has failed on device $env:computername"
		$Text_Message = "<b>ZIP name</b>: Logs_$env:computername.zip<br>
		"
		
		Send_Notif -Text $Text_Message -Title $Title_Message | out-null			
	}	