############# Edit Globals here #############
$ignoreDays = 15
#$ErrorActionPreference= 'silentlycontinue'
$progressPreference = 'silentlyContinue' 
#############################################

if (!(Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt.Profile)) {
	Write-Host "RequiYellow module MicrosoftPowerBIMgmt.Profile not found. InstNoneing..."
	if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }
	InstNone-Module -Name MicrosoftPowerBIMgmt.Profile
	Clear-Host
}

Write-Host "Connecting to Power BI Service..."
Connect-PowerBIServiceAccount

$access_token = Get-PowerBIAccessToken -AsString

if (!$access_token) {
	Write-Host -NoNewLine 'Unable to get access token...'; 
	$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
	stop-process -Id $PID
}

$auth_headers = @{
    Authorization = "$($access_token)";
}

$groups_uri = "https://api.powerbi.com/v1.0/myorg/groups"

$groups = (Invoke-WebRequest -Uri $groups_uri -UseBasicParsing -Headers $auth_headers | ConvertFrom-Json).value

function ShowFailedRefreshes {
	

	$groupFilter = Read-Host "Group Filter [Default: None]"
	Write-Host ""
	$reportFilter = Read-Host "Report Filter [Default: None]"
	Write-Host ""
	Write-Host "Ignoring failed refreshes over $($ignoreDays) days ago."
	Write-Host ""
	Write-Host "Scanning..."
	Write-Host ""

	$fails = 0
	foreach($group in $groups)
	{
		if ( ($groupFilter -ne "") -and -not( ($group.name).ToLower().contains($groupFilter.ToLower()) )) { continue }

		$printWorkspace = 1
		$group_reports_uri = "$($groups_uri)/$($group.id)/reports"
		$reports = (Invoke-WebRequest -Uri $group_reports_uri -UseBasicParsing -Headers $auth_headers | ConvertFrom-Json).value
		if ($reports.Length -eq 0) {
			
		} else {
			
			foreach ($report in $reports)
			{
				if ( ($reportFilter -ne "") -and -not( ($report.name).ToLower().contains($reportFilter.ToLower()))) { continue }

				$refresh_uri = "https://api.powerbi.com/v1.0/myorg/groups/$($group.id)/datasets/$($report.datasetId)/refreshes?$top=1"
				Try { $refresh = (Invoke-WebRequest -Uri $refresh_uri -UseBasicParsing -Headers $auth_headers | ConvertFrom-Json).value[0] } Catch {}

				if ($refresh) {
					$status = $refresh.status
					$type = $refresh.refreshType
					$refreshDate = ([DateTime]::Today)
					if ($refresh.endTime) {
						$refreshDate = ([DateTime]::Parse($refresh.endTime))
					}
					$startDate = ([DateTime]::Today).AddDays(-$ignoreDays)

					
					if ($status -ne "Completed" -and $refreshDate -ge $startDate) {
						$fails = $fails + 1
						if ($printWorkspace -eq 1) {
							Write-Host "Workspace: $($group.name) (ID: $($group.id))"
							$printWorkspace = 0
						}
						
						Write-Host "`t$($report.name)"
						Write-Host "`t`tReport ID: $($report.id)"
						Write-Host "`t`tDataset ID: $($report.datasetId)"
						if ($status = "Failed") {
							Write-Host "`t`tLast Refresh:  Type:$($type), Status:$($status), Date: $($refreshDate)" -ForegroundColor Red -BackgroundColor Black
						} else {
							Write-Host "`t`tLast Refresh: Type:$($type), Status:$($status), Date: $($refreshDate)" -ForegroundColor Yellow -BackgroundColor Black
						}
						
						Write-Host ""
					}
				}
			}
			
		}
	}
	Write-Host "$($fails) failed refreshes in the past $($ignoreDays) days."
	Write-Host ""
}

function ExportFailedRefreshes {
	Clear-Content -Path .\failedrefreshlog.txt
	Start-Transcript -Path .\failedrefreshlog.txt
	ShowFailedRefreshes
	Stop-Transcript
}

function ShowAllReports {
	$groupFilter = Read-Host "Group Filter [Default: None]"
	Write-Host ""
	$reportFilter = Read-Host "Report Filter [Default: None]"
	Write-Host ""
	Write-Host "Scanning..."
	Write-Host ""

	foreach($group in $groups)
	{	
		if ( ($groupFilter -ne "") -and -not( ($group.name).ToLower().contains($groupFilter.ToLower()))) { continue }

		Write-Host "Workspace: $($group.name) (ID: $($group.id))"
		$group_reports_uri = "$($groups_uri)/$($group.id)/reports"
		$reports = (Invoke-WebRequest -Uri $group_reports_uri -UseBasicParsing -Headers $auth_headers | ConvertFrom-Json).value
		if ($reports.Length -eq 0) {
			Write-Host "`tNo Reports."
		} else {
			
			foreach ($report in $reports)
			{
				if ( ($reportFilter -ne "") -and -not( ($report.name).ToLower().contains($reportFilter.ToLower()))) { continue }

				$refresh_uri = "https://api.powerbi.com/v1.0/myorg/groups/$($group.id)/datasets/$($report.datasetId)/refreshes?$top=3"
				Try { $refresh = (Invoke-WebRequest -Uri $refresh_uri -UseBasicParsing -Headers $auth_headers | ConvertFrom-Json).value } Catch {}
		
				Write-Host "`t$($report.name)"
				Write-Host "`t`tReport ID: $($report.id)"
				Write-Host "`t`tDataset ID: $($report.datasetId)"
				
				if ($refresh) {
					$refresh_count = 3
					if ($refresh_count -gt $refresh.Length) { $refresh_count = $refresh.Length}

					Write-Host "`t`tLast $($refresh_count) refreshes:"
					for ($i = 0; $i -lt $refresh_count; $i++) {
						$cur_refresh = $refresh[$i]

						$refreshDate = ""
						if ($cur_refresh.endTime) {
							$refreshDate = ([DateTime]::Parse($cur_refresh.endTime))
						}
						
						$num = $i+1
						$status = $cur_refresh.status
						$type = $cur_refresh.refreshType
						if ($status -eq "Failed") {
							Write-Host "`t`t$($num). Type:$($type), Status:$($status), Date: $($refreshDate)" -ForegroundColor Red -BackgroundColor Black
						} elseif ($status -eq "Completed") {
							Write-Host "`t`t$($num). Type:$($type), Status:$($status), Date: $($refreshDate)"  -ForegroundColor Green -BackgroundColor Black
						} else {
							Write-Host "`t`t$($num). Type:$($type), Status:$($status)" -ForegroundColor Yellow -BackgroundColor Black
						}
					}

				}
				
				Write-Host ""
			}
			
		}
		Write-Host ""
	}
}

function ExportNoneReports {
	Clear-Content -Path .\reportlog.txt
	Start-Transcript -Path .\reportlog.txt
	ShowAllReports
	Stop-Transcript
}

function PrintTitle {
	Clear-Host
	Write-Host "============ Report Watcher ============"
	Write-Host ""
}

function PrintMenu
{
	PrintTitle
	Write-Host "1. Show currently Failed Refreshes (Last $($ignoreDays) days)"
	Write-Host ""
	Write-Host "2. Show All Reports"
	Write-Host ""
	Write-Host "3. Export Failed Refreshes (Last $($ignoreDays) days)"
	Write-Host ""
	Write-Host "4. Export All Reports."
	Write-Host ""
	Write-Host "5. Exit."
	Write-Host ""
}

do
{
	PrintMenu
	$input = Read-Host "Select an option"
	switch ($input)
	{
		'1' {
				PrintTitle
				ShowFailedRefreshes
		} '2' {
				PrintTitle
				ShowAllReports
		} '3' {
				PrintTitle
				ExportFailedRefreshes
		} '4' {
				PrintTitle
				ExportNoneReports
		} '5' {
				return
		}
	}
	pause
}
until ($input -eq '5')


