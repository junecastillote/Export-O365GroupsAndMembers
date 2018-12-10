$scriptVersion = "1.0"
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
[xml]$config = Get-Content "$($script_root)\config.xml"
#debug-----------------------------------------------------------------------------------------------
[boolean]$enableDebug = $config.options.enableDebug
#----------------------------------------------------------------------------------------------------
#Mail------------------------------------------------------------------------------------------------
[boolean]$sendReport = $config.options.SendReport
[string]$tenantName = $config.options.tenantName
[string]$fromAddress = $config.options.fromAddress
[string]$toAddress = $config.options.toAddress
[string]$smtpServer = "smtp.office365.com"
[int]$smtpPort = "587"
[string]$mailSubject = "Office 365 Distribution Lists Backup"
#----------------------------------------------------------------------------------------------------
#Housekeeping----------------------------------------------------------------------------------------
$enableHousekeeping = $true #Set this to false if you do not want old backups to be deleted
$daysToKeep = 60
#----------------------------------------------------------------------------------------------------
$Today=Get-Date
[string]$filePrefix = '{0:dd-MMM-yyyy_hh-mm_tt}' -f $Today
$logFile = "$($script_root)\Logs\DebugLog_$($filePrefix).txt"
$BackupPath = "$($script_root)\BackupDir"
$BackupFile="$($BackupPath)\Export_$($filePrefix).csv"
$zipFile="$($BackupPath)\Export_$($filePrefix).zip"
#Functions------------------------------------------------------------------------------------------
#Function to connect to EXO Shell
Function New-EXOSession
{
    [CmdletBinding()]
    param(
        [parameter(mandatory=$true)]
        [PSCredential] $exoCredential
    )

    Get-PSSession | Remove-PSSession -Confirm:$false
    $EXOSession = New-PSSession -ConfigurationName "Microsoft.Exchange" -ConnectionUri 'https://ps.outlook.com/powershell' -Credential $exoCredential -Authentication Basic -AllowRedirection
    #$office365 = Import-PSSession $EXOSession -AllowClobber -DisableNameChecking
    Import-PSSession $EXOSession -AllowClobber -DisableNameChecking | out-null
}

#Function to compress the CSV file
Function New-ZipFile
{
	[CmdletBinding()] 
    param ( 
        [Parameter(Mandatory)] 
        [string]$fileToZip,
    
		[Parameter(Mandatory)]
		[string]$destinationZip
	)
	Add-Type -assembly System.IO.Compression
	Add-Type -assembly System.IO.Compression.FileSystem
	[System.IO.Compression.ZipArchive]$outZipFile = [System.IO.Compression.ZipFile]::Open($destinationZip, ([System.IO.Compression.ZipArchiveMode]::Create))
	[System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($outZipFile, $fileToZip, (Split-Path $fileToZip -Leaf)) | out-null
	$outZipFile.Dispose()
}

#Function to delete old files based on age
Function Invoke-Housekeeping
{
    [CmdletBinding()] 
    param ( 
        [Parameter(Mandatory)] 
        [string]$folderPath,
    
		[Parameter(Mandatory)]
		[int]$daysToKeep
    )
    
    $datetoDelete = (Get-Date).AddDays(-$daysToKeep)
    $filesToDelete = Get-ChildItem $FolderPath | Where-Object { $_.LastWriteTime -lt $datetoDelete }

    if (($filesToDelete.Count) -gt 0) {	
		foreach ($file in $filesToDelete) {
            Remove-Item -Path ($file.FullName) -Force -ErrorAction SilentlyContinue
		}
	}	
}
#----------------------------------------------------------------------------------------------------
#kill transcript if still running--------------------------------------------------------------------
try{
    stop-transcript|out-null
  }
  catch [System.InvalidOperationException]{}
#----------------------------------------------------------------------------------------------------
#start transcribing----------------------------------------------------------------------------------
if ($enableDebug -eq $true) {Start-Transcript -Path $logFile}
#----------------------------------------------------------------------------------------------------
#BEGIN------------------------------------------------------------------------------------------
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Begin" -ForegroundColor Green
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Connecting to Exchange Online Shell" -ForegroundColor Green

#Connect to O365 Shell
#Note: This uses an encrypted credential (XML). To store the credential:
#1. Login to the Server/Computer using the account that will be used to run the script/task
#2. Run this "Get-Credential | Export-CliXml ExOnlineStoredCredential.xml"
#3. Make sure that ExOnlineStoredCredential.xml is in the same folder as the script.
$onLineCredential = Import-Clixml "$($script_root)\ExOnlineStoredCredential.xml"
New-EXOSession $onLineCredential

#Start Export Process---------------------------------------------------------------------------
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Retrieving Distribution Groups" -ForegroundColor Yellow
$grouplist = Get-DistributionGroup -ResultSize Unlimited | Sort-Object Name
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": There are a total of $($grouplist.count) groups" -ForegroundColor Yellow

$i=1
foreach ($group in $grouplist)	{
	$temp = "" | Select-Object Name,Identity,SamAccountName,GroupType,BypassNestedModerationEnabled, `
	ManagedBy,MemberJoinRestriction,MemberDepartRestriction,ReportToManagerEnabled, `
	ReportToOriginatorEnabled,SendOofMessageToOriginatorEnabled,AcceptMessagesOnlyFrom, `
	AcceptMessagesOnlyFromDLMembers,AcceptMessagesOnlyFromSendersOrMembers,Alias,OrganizationalUnit, `
	DisplayName,EmailAddresses,GrantSendOnBehalfTo,HiddenFromAddressListsEnabled,MaxSendSize,MaxReceiveSize, `
	ModeratedBy,ModerationEnabled,EmailAddressPolicyEnabled,PrimarySmtpAddress,RecipientType, `
	RecipientTypeDetails,RejectMessagesFrom,RejectMessagesFromDLMembers,RejectMessagesFromSendersOrMembers, `
	RequireSenderAuthenticationEnabled,SendModerationNotifications,MailTip,MembersUPN

	[array]$group_members = Get-DistributionGroupMember -id $group.DistinguishedName -ResultSize Unlimited
	Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ($($i) of $($grouplist.count)) | [$($group_members.count) member/s] | $($group.DisplayName)" -ForegroundColor Yellow
	$temp.Name = $group.Name
	$temp.Identity = $group.Identity
	$temp.SamAccountName = $group.SamAccountName
	$temp.GroupType = $group.GroupType
	$temp.BypassNestedModerationEnabled = $group.BypassNestedModerationEnabled
	if ($group.ManagedBy){$temp.ManagedBy = [string]::join(";", ($group.ManagedBy.Split(",")))}			
	$temp.MemberJoinRestriction = $group.MemberJoinRestriction
	$temp.MemberDepartRestriction = $group.MemberDepartRestriction
	$temp.ReportToManagerEnabled = $group.ReportToManagerEnabled
	$temp.ReportToOriginatorEnabled = $group.ReportToOriginatorEnabled
	$temp.SendOofMessageToOriginatorEnabled = $group.SendOofMessageToOriginatorEnabled
	if ($group.AcceptMessagesOnlyFrom) {$temp.AcceptMessagesOnlyFrom = [string]::join(";", ($group.AcceptMessagesOnlyFrom.Split(",")))}
	if ($group.AcceptMessagesOnlyFromDLMembers) {$temp.AcceptMessagesOnlyFromDLMembers = [string]::join(";", ($group.AcceptMessagesOnlyFromDLMembers.Split(",")))}
	if ($group.AcceptMessagesOnlyFromSendersOrMembers) {$temp.AcceptMessagesOnlyFromSendersOrMembers = [string]::join(";", ($group.AcceptMessagesOnlyFromSendersOrMembers.Split(",")))}			
	$temp.Alias = $group.Alias
	$temp.OrganizationalUnit = $group.OrganizationalUnit
	$temp.DisplayName = $group.DisplayName
	if ($group.EmailAddresses) {$temp.EmailAddresses = [string]::join(";", ($group.EmailAddresses.Split(",")))}
	if ($group.GrantSendOnBehalfTo) {$temp.GrantSendOnBehalfTo = [string]::join(";", ($group.GrantSendOnBehalfTo.Split(",")))}			
	$temp.HiddenFromAddressListsEnabled = $group.HiddenFromAddressListsEnabled
	$temp.MaxSendSize = $group.MaxSendSize
	$temp.MaxReceiveSize = $group.MaxReceiveSize
	if ($group.ModeratedBy) {$temp.ModeratedBy = [string]::join(";", ($group.ModeratedBy.Split(",")))}			
	$temp.ModerationEnabled = $group.ModerationEnabled
	$temp.EmailAddressPolicyEnabled = $group.EmailAddressPolicyEnabled
	$temp.PrimarySmtpAddress = $group.PrimarySmtpAddress
	$temp.RecipientType = $group.RecipientType
	$temp.RecipientTypeDetails = $group.RecipientTypeDetails
	if ($group.RejectMessagesFrom) {$temp.RejectMessagesFrom = [string]::join(";", ($group.RejectMessagesFrom.Split(",")))}
	if ($group.RejectMessagesFromDLMembers) {$temp.RejectMessagesFromDLMembers = [string]::join(";", ($group.RejectMessagesFromDLMembers.Split(",")))}
	if ($group.RejectMessagesFromSendersOrMembers) {$temp.RejectMessagesFromSendersOrMembers = [string]::join(";", ($group.RejectMessagesFromSendersOrMembers.Splt(",")))}			
	$temp.RequireSenderAuthenticationEnabled = $group.RequireSenderAuthenticationEnabled
	$temp.SendModerationNotifications = $group.SendModerationNotifications
	$temp.MailTip = $group.MailTip
	if ($group_members.count -gt 0)	{ $temp.MembersUPN = $group_members.WindowsLiveID -join ";"	}
	$temp | Export-Csv $BackupFile -NoTypeInformation -Append		
	$i=$i+1
}
#-----------------------------------------------------------------------------------------------
#Zip the file to save space---------------------------------------------------------------------
#for PS v5+
#Compress-Archive -LiteralPath $BackupFile -CompressionLevel Optimal -DestinationPath $zipFile
#for PS v4
New-ZipFile -fileToZip $BackupFile -destinationZip $zipFile
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Backup Saved to $zipFile" -ForegroundColor Yellow
$zipSize = (Get-ChildItem $zipFile | Measure-Object -Property Length -Sum)
#Allow some time (in seconds) for the file access to close, increase especially if the resulting files are huge, or server I/O is busy.
$sleepTime=5
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Pending next operation for $($sleepTime) seconds" -ForegroundColor Yellow
Start-Sleep -Seconds $sleepTime
Remove-Item $BackupFile
#-----------------------------------------------------------------------------------------------
#Invoke Housekeeping----------------------------------------------------------------------------
if ($enableHousekeeping -eq $true){
	Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Deleting backup files older than $($daysToKeep) days" -ForegroundColor Yellow
	Invoke-Housekeeping -folderPath $BackupPath -daysToKeep $daysToKeep
}
#-----------------------------------------------------------------------------------------------
#Count the number of backups existing and the total size----------------------------------------
$BackupCount = (Get-ChildItem $BackupPath -recurse | Measure-Object -Property Length -Sum)
#-----------------------------------------------------------------------------------------------
$timeTaken = New-TimeSpan -Start $Today -End (Get-Date)
#Send email if option is enabled ---------------------------------------------------------------
if ($SendReport -eq $true){	
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Sending report to" ($toAddress -join ";") -ForegroundColor Yellow
$xSubject="[$($tenantName)] $($mailSubject): " + ('{0:dd-MMM-yyyy hh:mm:ss tt}' -f $Today)
$htmlBody=@'
<!DOCTYPE html>
<html>
<head>
<style>
table {
  font-family: "Century Gothic", sans-serif;
  border-collapse: collapse;
  width: 100%;
}
td, th {
  border: 1px solid #dddddd;
  text-align: left;
  padding: 8px;
}

</style>
</head>
<body>
<table>
'@
$htmlBody+="<tr><th>----SUMMARY----</th></tr>"
$htmlBody+="<tr><th>Number of Groups</th><td>$($grouplist.count)</td></tr>"
$htmlBody+="<tr><th>Backup ServerServer</th><td>"+(Get-Content env:computername)+"</td></tr>"
$htmlBody+="<tr><th>Backup File</th><td>$($zipFile)</td></tr>"
$htmlBody+="<tr><th>Backup Size</th><td>"+ ("{0:N2}" -f ($zipSize.Sum / 1KB)) + " KB</td></tr>"
$htmlBody+="<tr><th>Time to Complete</th><td>"+ ("{0:N2}" -f $($timeTaken.TotalMinutes)) + " Minutes</td></tr>"
$htmlBody+="<tr><th>Total Number of Backups</th><td>$($BackupCount.Count)</td></tr>"
$htmlBody+="<tr><th>Total Backup Folder Size</th><td>"+ ("{0:N2}" -f ($BackupCount.Sum / 1KB)) + " KB</td></tr>"
$htmlBody+="<tr><th>----SETTINGS----</th></tr>"
$htmlBody+="<tr><th>Tenant Organization</th><td>$($tenantName)</td></tr>"
$htmlBody+="<tr><th>Debug Enabled</th><td>$($enableDebug)</td></tr>"
$htmlBody+="<tr><th>Housekeeping Enabled</th><td>$($enableHousekeeping)</td></tr>"
$htmlBody+="<tr><th>Days to Keep</th><td>$($daysToKeep)</td></tr>"
$htmlBody+="<tr><th>Report Recipients</th><td>" + $toAddress.Replace(",","<br>") + "</td></tr>"
$htmlBody+="<tr><th>SMTP Server</th><td>$($smtpServer)</td></tr>"
$htmlBody+="<tr><th>Script Path</th><td>$($MyInvocation.MyCommand.Definition)</td></tr>"
$htmlBody+="<tr><th>Script Source Site</th><td><a href=""https://github.com/junecastillote/Export-O365GroupsAndMembers"">Export-O365GroupsAndMembers.ps1</a> version $($scriptVersion)</td></tr>"
$htmlBody+="</table></body></html>"
Send-MailMessage -from $fromAddress -to $toAddress.Split(",") -subject $xSubject -body $htmlBody -dno onSuccess, onFailure -smtpServer $SMTPServer -Port $smtpPort -Credential $onLineCredential -UseSsl -BodyAsHtml
}
#-----------------------------------------------------------------------------------------------
Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": End" -ForegroundColor Green
#-----------------------------------------------------------------------------------------------
#kill transcript if still running
try{
    stop-transcript|out-null
  }
  catch [System.InvalidOperationException]{}