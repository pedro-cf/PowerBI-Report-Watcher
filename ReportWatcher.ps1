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
	

	$groupFilter = Read-Host "Group Filter"
	Write-Host ""
	$reportFilter = Read-Host "Report Filter"
	Write-Host ""
	Write-Host "Ignoring failed refreshes over $($ignoreDays) days ago."
	Write-Host ""
	Write-Host "Printing failed refreshes..."
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
					} elseif ($refresh.startTime) {
						$refreshDate = ([DateTime]::Parse($refresh.startTime))
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
						if ($status -eq "Failed") {
							Write-Host "`t`tLast Refresh:  Type: $($type), Status: $($status), Date: $($refreshDate)" -ForegroundColor White -BackgroundColor DarkRed
							
							$error_json_ex = $refresh.serviceExceptionJson 
							
							if ($error_json_ex) {
							
								$error_json = $error_json_ex | ConvertFrom-Json | ConvertTo-Json -depth 100 | ConvertFrom-Json
								
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
	Write-Host "Printing reports..."
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
						} elseif ($cur_refresh.startTime) {
							$refreshDate = ([DateTime]::Parse($cur_refresh.startTime))
						}
						
						$num = $i+1
						$status = $cur_refresh.status
						$type = $cur_refresh.refreshType
						if ($status -eq "Failed") {
							Write-Host "`t`t$($num). Type: $($type), Status: $($status), Date: $($refreshDate)" -ForegroundColor White -BackgroundColor DarkRed

							$error_json_ex = $cur_refresh.serviceExceptionJson
							
							if ($error_json_ex) {
							
								$error_json = $error_json_ex | ConvertFrom-Json | ConvertTo-Json -depth 100 | ConvertFrom-Json
								
								
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

function ExportAllReports {
	Clear-Content -Path .\reportlog.txt
	Start-Transcript -Path .\reportlog.txt
	ShowAllReports
	Stop-Transcript
}

function ShowAllScheduledRefreshes {
	$groupFilter = Read-Host "Group Filter"
	Write-Host ""
	$reportFilter = Read-Host "Report Filter"
	Write-Host ""
	Write-Host "Printing Scheduled Refreshes..."
	Write-Host ""
	Write-Host "UTC`tDuration`tDataset"
	Write-Host "-----`t--------`t-------"

	$schedule_list = @()

	foreach($group in $groups)
	{	
		if ( ($groupFilter -ne "") -and -not( ($group.name).ToLower().contains($groupFilter.ToLower()))) { continue }

		$group_reports_uri = "$($groups_uri)/$($group.id)/reports"
		$reports = (Invoke-WebRequest -Uri $group_reports_uri -UseBasicParsing -Headers $auth_headers | ConvertFrom-Json).value
		if ($reports.Length -eq 0) {

		} else {
			
			foreach ($report in $reports)
			{
				if ( ($reportFilter -ne "") -and -not( ($report.name).ToLower().contains($reportFilter.ToLower()))) { continue }
				
				$scheduled_refresh_uri = "https://api.powerbi.com/v1.0/myorg/groups/$($group.id)/datasets/$($report.datasetId)/refreshSchedule"
				Try { $scheduled_refresh = (Invoke-WebRequest -Uri $scheduled_refresh_uri -UseBasicParsing -Headers $auth_headers | ConvertFrom-Json) } Catch {}


				if ($scheduled_refresh) {
					if ($scheduled_refresh.enabled -eq "True") {

						$refresh_history_uri = "https://api.powerbi.com/v1.0/myorg/groups/$($group.id)/datasets/$($report.datasetId)/refreshes?$top=1"
						Try { $refresh_history = (Invoke-WebRequest -Uri $refresh_history_uri -UseBasicParsing -Headers $auth_headers | ConvertFrom-Json).value } Catch {}

						$refreshDuration =  "????????"
						if ($refresh_history) {
							$refresh_count = $refresh_history.Length
							if ($refresh_count -gt 0) {
								$cur_refresh = $refresh_history[0]
								if ($cur_refresh.endTime -and $cur_refresh.startTime) {
									$refreshStart = ([DateTime]::Parse($cur_refresh.endTime))
									$refreshEnd = ([DateTime]::Parse($cur_refresh.startTime))
									$refreshDuration = ("{0:hh\:mm\:ss}" -f ($refreshStart - $refreshEnd))
								}
							}
						}
					
						if ($scheduled_refresh.times)  {
							foreach($time in $scheduled_refresh.times) {
								$schedule_list += "$($time)`t$($refreshDuration)`t$($group.name) > $($report.name)"
							}
						} else {
							$schedule_list += "00:00`t$($refreshDuration)`t$($group.name) > $($report.name)"
						}
					}
				}
			}
		}
	}

	$time_segments = @(
		"00:00", "00:30", "01:00", "01:30", "02:00", "02:30", "03:00", "03:30", "04:00", "04:30",
		"05:00", "05:30", "06:00", "06:30", "07:00", "07:30", "08:00", "08:30", "09:00", "09:30",
		"10:00", "10:30", "11:00", "11:30","12:00", "12:30", "13:00", "13:30", "14:00", "14:30",
		"15:00","15:30", "16:00", "16:30", "17:00", "17:30", "18:00", "18:30", "19:00", "19:30",
		"20:00", "20:30", "21:00", "21:30", "22:00", "22:30", "23:00", "23:30"
	)

	foreach ($time_segment in $time_segments) {
		$hasOne = 0
		foreach($schedule_item in $schedule_list) {
			if ($schedule_item.StartsWith($time_segment)) {
				Write-Host $schedule_item
				$hasOne = 1
			}
		}
		if ($hasOne -eq 0) {
			Write-Host "$($time_segment)`t---"
		}
	}

	Write-Host ""
}

function ExportAllScheduledRefreshes {
	Clear-Content -Path .\scheduledRefreshLog.txt
	Start-Transcript -Path .\scheduledRefreshLog.txt
	ShowAllScheduledRefreshes
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
	Write-Host "3. Show All Scheduled Refreshes"
	Write-Host ""
	Write-Host "4. Export Failed Refreshes (Last $($ignoreDays) days)"
	Write-Host ""
	Write-Host "5. Export All Reports"
	Write-Host ""
	Write-Host "6. Export All Scheduled Refreshes"
	Write-Host ""
	Write-Host "7. Exit"
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
				ShowAllScheduledRefreshes
				
		} '4' {
				PrintTitle
				ExportFailedRefreshes
		} '5' {
				PrintTitle
				ExportAllReports
		} '6' {
				PrintTitle
				ExportAllScheduledRefreshes
		} '7' {
				return
		}
	}
	pause
}
until ($input -eq '7')


