# Usage
# Expired_Users_Script

Param ([string]$fromList)

$date = Get-Date -format yyyy-MM-dd
$logFileDir = "C:\PSLogFiles"
$logFileSuccess = $logFileDir+"\list_"+$date+"_success.log"
$logFileError = $logFileDir+"\list__"+$date+"_error.log"

Write-Host "`r`nCreating log files for output in:" -foregroundcolor "Green"

New-Item $logFileSuccess -type file -force
New-Item $logFileError -type file -force

function LogWriteSuccess {
   Param ([string]$logString)
   Add-content $logFileSuccess -value $logString
}
function LogWriteError {
   Param ([string]$logString)
   Add-content $logFileError -value $logString
}

$ListName = ['List Name']
$WorkflowName = 'Expiration Notice'

$web = get-spweb ['website']
$List = $Web.Lists[$ListName]
$Items = $List.GetItems()

$Manager = $Web.Site.WorkflowManager
$Association = $List.WorkflowAssociations.GetAssociationByName($WorkflowName, 'en-US')
$Data = $Association.AssociationData

Write-Host "`r`nImporting [your list name] " $list.Title'...' -foregroundcolor "Green"
Write-Host "Checking [list date]..." -foregroundcolor "Green"
$ActiveUserCount = 0
$ExpiredUserCount = 0

$DateToCompare = [DateTime]::Now

foreach ($item in $Items | where {$_["Active"] -eq $TRUE}) {
	$quizExpireDate = $item["QuizDate"].AddDays(+365)
	$ActiveUserCount++
	if ($DateToCompare -gt $quizExpireDate) {
		$logWriteString = "Quiz certification for " + $item["email"] + " expired on " + $quizExpireDate.ToString("MM/dd/yyyy")
		LogWriteSuccess $logWriteString
		$item["triggerWorkflow"] = "Yes"
		$item.Update()
		$WF = $Manager.StartWorkflow($Item, $Association, $Data, $true)
		$ExpiredUserCount++
	}
	else {
		# Do Nothing
		# Write-Host "Quiz date is still in range"
		# Write-Host $quizExpireDate ' > ' $todayDate
	}	
}

Write-Host "`r`nActive user count:" $ActiveUserCount -foregroundcolor "Green"
LogWriteSuccess "`r`nActive user count:" $ActiveUserCount
Write-Host "Expired user count:" $ExpiredUserCount -foregroundcolor "Green"
LogWriteSuccess "Expired user count:" $ExpiredUserCount
#Write-Host "Finished ForEach loop..." -foregroundcolor "Green"
Write-Host "`r`nProcess finished, updating web object..." -foregroundcolor "Green"

#Open Workflow Summary List
#Write generic values to list
#Trigger workflow on list

$web.Update()

Write-Host "Log files can be found in C:\PSLogFiles" -foregroundcolor "Green"

$web.Dispose()
