# PowerBI Report Watcher

PowerBI Report Watcher is a PowerShell script that provides various functionalities to monitor and manage Power BI reports.

## Usage

To use the script, you can follow these steps:
1. Open PowerShell and navigate to the directory where the script is located.
2. Run the script using the following command:
    ```
    .\PowerBIReportWatcher.ps1
    ```
3. The script will ask for your Power BI credentials. Enter your email and password, and then select the environment you want to work with.
4. Once you're authenticated, you can select from the following options:
    1. Show currently Failed Refreshes (Last ***$ignoreDays*** days)
    2. Show All Reports
    3. Show All Scheduled Refreshes
    4. Export Failed Refreshes (Last ***$ignoreDays*** days)
    5. Export All Reports
    6. Export All Scheduled Refreshes

## Requirements

PowerBI Report Watcher requires the following:
- Windows PowerShell 5.1 or later
- Azure Active Directory (AD) module for Windows PowerShell
- Microsoft Power BI module for Windows PowerShell

You can install the required modules using the following commands:

```
Install-Module AzureAD
Install-Module -Name MicrosoftPowerBIMgmt
```



## License

PowerBI Report Watcher is released under the MIT License. See the [LICENSE](LICENSE) file for details.

## Support

If you find this project useful, consider buying me a coffee! Donations help keep this project going and are greatly appreciated.

[![Buy me a coffee](https://img.shields.io/badge/-Buy%20me%20a%20coffee-orange?logo=buy-me-a-coffee&logoColor=white&style=for-the-badge)](https://www.buymeacoffee.com/pedro_cf)

