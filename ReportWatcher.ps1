$ErrorActionPreference= 'silentlycontinue'

if (!(Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt.Profile)) {
	Write-Host "Required module MicrosoftPowerBIMgmt.Profile not found. Installing..."
	if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }
	Install-Module -Name MicrosoftPowerBIMgmt.Profile
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
	foreach($group in $groups)
	{
		$printWorkspace = 1
		$group_reports_uri = "$($groups_uri)/$($group.id)/reports"
		$reports = (Invoke-WebRequest -Uri $group_reports_uri -UseBasicParsing -Headers $auth_headers | ConvertFrom-Json).value
		if ($reports.Length -eq 0) {
			
		} else {
			
			foreach ($report in $reports)
			{
				$refresh_uri = "https://api.powerbi.com/v1.0/myorg/groups/$($group.id)/datasets/$($report.datasetId)/refreshes?$top=1"
				Try { $refresh = (Invoke-WebRequest -Uri $refresh_uri -UseBasicParsing -Headers $auth_headers | ConvertFrom-Json).value[0] } Catch {}

				if ($refresh) {
					if ($refresh.refreshType -eq "Scheduled") {
						$status = $refresh.status
						if ($status -eq "Failed") {
							if ($printWorkspace -eq 1) {
								Write-Host "Workspace: $($group.name) (ID: $($group.id))" -ForegroundColor Red -BackgroundColor Black
								$printWorkspace = 0
							}
							
							Write-Host "`t$($report.name)" -ForegroundColor Red -BackgroundColor Black
							Write-Host "`t`tReport ID: $($report.id)" -ForegroundColor Red -BackgroundColor Black
							Write-Host "`t`tDataset ID: $($report.datasetId)" -ForegroundColor Red -BackgroundColor Black
							Write-Host "`t`tLast Refresh Date: $($refresh.endTime)" -ForegroundColor Red -BackgroundColor Black
							Write-Host "`t`tLast Refresh Status: $($refresh.status)" -ForegroundColor Red -BackgroundColor Black
							Write-Host ""
						}
					}
				}
			}
			
		}
	}
}

function ExportFailedRefreshes {
	Clear-Content -Path .\failedrefreshlog.txt
	Start-Transcript -Path .\failedrefreshlog.txt
	ShowFailedRefreshes
	Stop-Transcript
}

function ShowAllReports {
	foreach($group in $groups)
	{	
		Write-Host "Workspace: $($group.name) (ID: $($group.id))"
		$group_reports_uri = "$($groups_uri)/$($group.id)/reports"
		$reports = (Invoke-WebRequest -Uri $group_reports_uri -UseBasicParsing -Headers $auth_headers | ConvertFrom-Json).value
		if ($reports.Length -eq 0) {
			Write-Host "`tNo Reports."
		} else {
			
			foreach ($report in $reports)
			{
				$refresh_uri = "https://api.powerbi.com/v1.0/myorg/groups/$($group.id)/datasets/$($report.datasetId)/refreshes?$top=1"
				Try { $refresh = (Invoke-WebRequest -Uri $refresh_uri -UseBasicParsing -Headers $auth_headers | ConvertFrom-Json).value[0] } Catch {}
		
				Write-Host "`t$($report.name)"
				Write-Host "`t`tReport ID: $($report.id)"
				Write-Host "`t`tDataset ID: $($report.datasetId)"
				
				if ($refresh) {
					if ($refresh.refreshType -eq "Scheduled") {
						$status = $refresh.status
						if ($status -eq "Failed") {
							Write-Host "`t`tLast Refresh Date: $($refresh.endTime)" -ForegroundColor Red -BackgroundColor Black
							Write-Host "`t`tLast Refresh Status: $($refresh.status)" -ForegroundColor Red -BackgroundColor Black
						} else {
							Write-Host "`t`tLast Refresh Date: $($refresh.endTime)"
							Write-Host "`t`tLast Refresh Status: $($refresh.status)"
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

function PrintMenu
{
	cls
	Write-Host "====== Report Watcher ======"
	Write-Host ""
	Write-Host "1: Show Failed Refreshes."
	Write-Host "2: Show All Reports."
	Write-Host "3: Export Failed Refreshes"
	Write-Host "4: Export All Reports."
	Write-Host "5: Exit."
	Write-Host ""
}

do
{
	PrintMenu
	$input = Read-Host "Select an option: "
	switch ($input)
	{
		'1' {
				cls
				ShowFailedRefreshes
		} '2' {
				cls
				ShowAllReports
		} '3' {
				cls
				ExportFailedRefreshes
		} '4' {
				cls
				ExportAllReports
		} '5' {
				return
		}
	}
	pause
}
until ($input -eq '5')


