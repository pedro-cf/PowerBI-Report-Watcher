############# Edit Globals here #############
$ignoreDays = 15
#$ErrorActionPreference= 'silentlycontinue'
$progressPreference = 'silentlyContinue' 
#############################################

if (!(Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt.Profile)) {
	Write-Host "RequiDarkRed module MicrosoftPowerBIMgmt.Profile not found. InstNoneing..."
	if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }
	InstNone-Module -Name MicrosoftPowerBIMgmt.Profile
	Clear-Host
}

Write-Host "Connecting to Power BI Service..."
#Connect-PowerBIServiceAccount
#$access_token = Get-PowerBIAccessToken -AsString
$access_token = "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ik4tbEMwbi05REFMcXdodUhZbkhRNjNHZUNYYyIsImtpZCI6Ik4tbEMwbi05REFMcXdodUhZbkhRNjNHZUNYYyJ9.eyJhdWQiOiJodHRwczovL2FuYWx5c2lzLndpbmRvd3MubmV0L3Bvd2VyYmkvYXBpIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvN2U1ZjllMGUtNzllNi00NDI3LTg2MjctYmE5YTVmN2NiNTFmLyIsImlhdCI6MTU1MjIzNTE1OCwibmJmIjoxNTUyMjM1MTU4LCJleHAiOjE1NTIyMzkwNTgsImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJBVlFBcS84S0FBQUFzcDFyUmxMcmtPMXZqMTlaWldETDRZRG10anNlZWtLOTVyaEJzbHpVU1FqZlJITWd0WktQTmNYT0paTUthYm1YM1hNRytoeDdHeHRTYW16RkNBY0VKWEx0ZGwwR0ZQQjRubmlhV0UweXFZND0iLCJhbXIiOlsicHdkIiwibWZhIl0sImFwcGlkIjoiZWEwNjE2YmEtNjM4Yi00ZGY1LTk1YjktNjM2NjU5YWU1MTIxIiwiYXBwaWRhY3IiOiIwIiwiZmFtaWx5X25hbWUiOiJGZXJuYW5kZXMiLCJnaXZlbl9uYW1lIjoiUGVkcm8iLCJpcGFkZHIiOiIyMTMuMjIuMjQwLjExNyIsIm5hbWUiOiJQZWRybyBGZXJuYW5kZXMiLCJvaWQiOiI4ZjY5Zjk2NC0zOTcwLTQ4ZmUtYWQ4NS1iMzFlMDBmMWUzZmIiLCJwdWlkIjoiMTAwMzIwMDAzODE3Q0IzMCIsInNjcCI6InVzZXJfaW1wZXJzb25hdGlvbiIsInN1YiI6IkQ3VlMtLVJxN2RwU2xnNGt4Y1NvUGtTOFBFb05lZ25kbFZXMFoxQno5TWciLCJ0aWQiOiI3ZTVmOWUwZS03OWU2LTQ0MjctODYyNy1iYTlhNWY3Y2I1MWYiLCJ1bmlxdWVfbmFtZSI6InBlZHJvLmZlcm5hbmRlc0Bpbm92cmV0YWlsLmNvbSIsInVwbiI6InBlZHJvLmZlcm5hbmRlc0Bpbm92cmV0YWlsLmNvbSIsInV0aSI6Im5fV2d1aXpaYWtpVUtweEVhM2hWQUEiLCJ2ZXIiOiIxLjAifQ.bIwE-U1apXG_9_WOi7QnxM9XxikX7qAmElQU0UbMa7dkG24HK3PdnHhFcbqDxODI9vseeaEclMOHQNAcXVd1EAQkzHB7E9S9x0Z-JzDQK1glDslfh6jKQdb3G9PhUi80UoJGBU244Xkkn04LIWJ_IyCuevWgufp0SbP3445ZZOkS9lazmh7u77YC35FKFFZ5oHS2R6OVoNQ3XcGmYU3Vlkk49ofiVfaERws2QoKYueqNn4vsT8mAZkeJNOUU8ut3Jl0ePQQBZIPL583xPw6XP919_CwP6UqUQeerDmrbsKtbkfpfIaY8LbyKkyet_kh0TVAAGQzg6fhUlyeNmsAU2g"


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
	

	$groupFilter = Read-Host "Group Filter"
	Write-Host ""
	$reportFilter = Read-Host "Report Filter"
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
							Write-Host "`t`tLast Refresh:  Type: $($type), Status: $($status), Date: $($refreshDate)" -ForegroundColor White -BackgroundColor DarkRed
							
							$error_json = $refresh.serviceExceptionJson | ConvertFrom-Json | ConvertTo-Json -depth 100 | ConvertFrom-Json
							$cnt = 0
							foreach($pe in $error_json.error."pbi.error"."parameters") {
								if ($cnt -eq 0) { Write-Host "`t`t   *Table errors: " -NoNewline -ForegroundColor Black -BackgroundColor DarkRed }
								if ($cnt++ -gt 0) { 	Write-Host " | " -NoNewline -ForegroundColor Black -BackgroundColor DarkRed }
								Write-Host "[$($pe.Value)]" -NoNewline -ForegroundColor Black -BackgroundColor DarkRed
							}
							if ($cnt -gt 0) { Write-Host ""}
							$cnt = 0
							$error_details = ""
							foreach($de in $error_json.error."pbi.error"."details") {
								if ($de.detail.value.Length -gt 2) {
									if ($cnt++ -gt 0) { $error_details = "$($error_details)  | " }
									$error_details = "$($error_details) $($de.detail.value) "
								}
							}
							if ($cnt -gt 0) {
								$error_details = $error_details.replace("`n","").replace("`r","")
								Write-Host "`t`t   *Error details: $($error_details)" -ForegroundColor Black -BackgroundColor DarkRed
							}

						} else {
							Write-Host "`t`tLast Refresh: Type: $($type), Status: $($status), Date: $($refreshDate)" -ForegroundColor White -BackgroundColor DarkRed
						}
						
						Write-Host ""
					}
				}
			}
			
		}
	}
	Write-Host ""
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
	$groupFilter = Read-Host "Group Filter"
	Write-Host ""
	$reportFilter = Read-Host "Report Filter"
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
							Write-Host "`t`t$($num). Type: $($type), Status: $($status), Date: $($refreshDate)" -ForegroundColor White -BackgroundColor DarkRed

							$error_json = $cur_refresh.serviceExceptionJson | ConvertFrom-Json | ConvertTo-Json -depth 100 | ConvertFrom-Json
							$cnt = 0
							foreach($pe in $error_json.error."pbi.error"."parameters") {
								if ($cnt -eq 0) { Write-Host "`t`t   *Table errors: " -NoNewline -ForegroundColor Black -BackgroundColor DarkRed }
								if ($cnt++ -gt 0) { 	Write-Host " | " -NoNewline -ForegroundColor Black -BackgroundColor DarkRed }
								Write-Host "[$($pe.Value)]" -NoNewline -ForegroundColor Black -BackgroundColor DarkRed
							}
							if ($cnt -gt 0) { Write-Host ""}
							$cnt = 0
							$error_details = ""
							foreach($de in $error_json.error."pbi.error"."details") {
								if ($de.detail.value.Length -gt 2) {
									if ($cnt++ -gt 0) { $error_details = "$($error_details)  | " }
									$error_details = "$($error_details) $($de.detail.value) "
								}
							}
							if ($cnt -gt 0) {
								$error_details = $error_details.replace("`n","").replace("`r","")
								Write-Host "`t`t   *Error details: $($error_details)" -ForegroundColor Black -BackgroundColor DarkRed
							}

						} elseif ($status -eq "Completed") {
							Write-Host "`t`t$($num). Type: $($type), Status: $($status), Date: $($refreshDate)"  -ForegroundColor White -BackgroundColor DarkGreen
						} else {
							Write-Host "`t`t$($num). Type: $($type), Status: $($status)" -ForegroundColor White -BackgroundColor DarkRed
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


