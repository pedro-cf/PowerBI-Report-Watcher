########## Edit Globals here ##########
$ignoreDays = 15
$ErrorActionPreference= 'silentlycontinue'
$progressPreference = 'silentlyContinue' 
#######################################

if (!(Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt.Profile)) {
	Write-Host "RequiYellow module MicrosoftPowerBIMgmt.Profile not found. InstNoneing..."
	if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }
	InstNone-Module -Name MicrosoftPowerBIMgmt.Profile
	Clear-Host
}

Write-Host "Connecting to Power BI Service..."
#Connect-PowerBIServiceAccount

#$access_token = Get-PowerBIAccessToken -AsString

$access_token = "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ik4tbEMwbi05REFMcXdodUhZbkhRNjNHZUNYYyIsImtpZCI6Ik4tbEMwbi05REFMcXdodUhZbkhRNjNHZUNYYyJ9.eyJhdWQiOiJodHRwczovL2FuYWx5c2lzLndpbmRvd3MubmV0L3Bvd2VyYmkvYXBpIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvN2U1ZjllMGUtNzllNi00NDI3LTg2MjctYmE5YTVmN2NiNTFmLyIsImlhdCI6MTU1MTk5MTg4OSwibmJmIjoxNTUxOTkxODg5LCJleHAiOjE1NTE5OTU3ODksImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJBVlFBcS84S0FBQUFyVS95Y0R1OXNUa0p3NmFPVy90YlVGSFU0WDRtQ2dFeVJCYmVIVXNnQmlRSGZudk9PWkJpWTdyQUljRGZwNm1RL2x0MlFORmtvMnNwMElEaDRUbEFzd3dtTFEzUmIvV2NHZHhEbGcwbmdtUT0iLCJhbXIiOlsicHdkIiwibWZhIl0sImFwcGlkIjoiZWEwNjE2YmEtNjM4Yi00ZGY1LTk1YjktNjM2NjU5YWU1MTIxIiwiYXBwaWRhY3IiOiIwIiwiZmFtaWx5X25hbWUiOiJGZXJuYW5kZXMiLCJnaXZlbl9uYW1lIjoiUGVkcm8iLCJpcGFkZHIiOiIyMTMuMjIuMjQwLjExNyIsIm5hbWUiOiJQZWRybyBGZXJuYW5kZXMiLCJvaWQiOiI4ZjY5Zjk2NC0zOTcwLTQ4ZmUtYWQ4NS1iMzFlMDBmMWUzZmIiLCJwdWlkIjoiMTAwMzIwMDAzODE3Q0IzMCIsInNjcCI6InVzZXJfaW1wZXJzb25hdGlvbiIsInN1YiI6IkQ3VlMtLVJxN2RwU2xnNGt4Y1NvUGtTOFBFb05lZ25kbFZXMFoxQno5TWciLCJ0aWQiOiI3ZTVmOWUwZS03OWU2LTQ0MjctODYyNy1iYTlhNWY3Y2I1MWYiLCJ1bmlxdWVfbmFtZSI6InBlZHJvLmZlcm5hbmRlc0Bpbm92cmV0YWlsLmNvbSIsInVwbiI6InBlZHJvLmZlcm5hbmRlc0Bpbm92cmV0YWlsLmNvbSIsInV0aSI6IlNJenVhUWtuaWtLajFMZ05LZk1CQUEiLCJ2ZXIiOiIxLjAifQ.AIfOG-N4CVKpIML7qDUhVe6cT_1WGj3aIqJKZbbZpmpntZwA_yh0r0wLyM_E-4zyXILYgdsz6PWAcCpM9CkbSbZMK94ysRe4PTDo11wxAiRlbNh_BlfIP8Xioc2omxy0s-Bw3kq8dsPbS7isd1sRCShmfBrcCRL9pFyOsNBuhZEnyCO3D9L099N92Fv8YGT0UO2FYXaTBNOvj6mnk85ruK_cmkQDyvhePMrzSlxPyun7rksfXu3n9le_g0f1hMprMrApTTMtQeTaDNKh4fjGjD9Ofn4KpRTXGdSE6zc6IjWt_7iotp5HtHtypS0PUx5RzSYa9pO_ZTyagCBHjhzJTQ"


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
					if ($refresh.refreshType -eq "Scheduled") {
						$status = $refresh.status
						$refreshDate = ([DateTime]::Parse($refresh.endTime))
						$startDate = ([DateTime]::Today).AddYears($ignoreDays)

						if ($status -eq "Failed" -and $refreshDate -ge $startDate) {
							$fails = $fails + 1
							if ($printWorkspace -eq 1) {
								Write-Host "Workspace: $($group.name) (ID: $($group.id))"
								$printWorkspace = 0
							}
							
							Write-Host "`t$($report.name)"
							Write-Host "`t`tReport ID: $($report.id)"
							Write-Host "`t`tDataset ID: $($report.datasetId)"
							Write-Host "`t`tLast Refresh Date: $($refreshDate)" -ForegroundColor Red -BackgroundColor Black
							Write-Host "`t`tLast Refresh Status: $($status)" -ForegroundColor Red -BackgroundColor Black
							Write-Host ""
						}
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

				$refresh_uri = "https://api.powerbi.com/v1.0/myorg/groups/$($group.id)/datasets/$($report.datasetId)/refreshes?$top=1"
				Try { $refresh = (Invoke-WebRequest -Uri $refresh_uri -UseBasicParsing -Headers $auth_headers | ConvertFrom-Json).value[0] } Catch {}
		
				Write-Host "`t$($report.name)"
				Write-Host "`t`tReport ID: $($report.id)"
				Write-Host "`t`tDataset ID: $($report.datasetId)"
				
				if ($refresh) {
					$refreshDate = ([DateTime]::Parse($refresh.endTime))

					if ($refresh.refreshType -eq "Scheduled") {
						$status = $refresh.status
						if ($status -eq "Failed") {
							Write-Host "`t`tLast Refresh Date: $($refreshDate)" -ForegroundColor Red -BackgroundColor Black
							Write-Host "`t`tLast Refresh Status: $($refresh.status)" -ForegroundColor Red -BackgroundColor Black
						} else {
							Write-Host "`t`tLast Refresh Date: $($refreshDate)" -ForegroundColor Green -BackgroundColor Black
							Write-Host "`t`tLast Refresh Status: $($refresh.status)" -ForegroundColor Green -BackgroundColor Black
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
	Write-Host "1 - Show Failed Refreshes (Last $($ignoreDays) days)"
	Write-Host ""
	Write-Host "2 - Show All Reports."
	Write-Host ""
	Write-Host "3 - Export Failed Refreshes (Last $($ignoreDays) days)"
	Write-Host ""
	Write-Host "4 - Export All Reports."
	Write-Host ""
	Write-Host "5 - Exit."
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


