# AD-Bulk-User-Import-Tool
A PowerShell GUI tool for analysing and importing bulk users into Active Directory from CSV files, even when the source data is inconsistently formatted.

## Features
- Accepts malformed or inconsistently formatted CSV files
- Analyses and prepares user data before import
- GUI-based selection of target OU
- GUI-based selection of security group
- Flexible username generation options
- Random or custom password generation
- Optional export of generated credentials on import of users
- Export option of all fields
- Written entirely in PowerShell

## Requirements
- Windows 10/11
- PowerShell 5.1 or later
- Domain Admin or delegated permissions

## How It Works

Right-click and Run with Powershell or use Powershell ISE or Visual Studio Code
1. Use the Browse button to the CSV file containing user data
2. Click on Load and view the data in the Preview Window
3. Use the Connect to AD button (if used on domain controller credentials are auto applied)
4. Use the Browse the OU button and select the OU
5. Click inside the Username box to select how you want usernames formatted
6. Click inside the Password box to select how you want the passwords formatted
7. Tick 'Assign Users to Security Groups' if required.
8. Then click on Select Groups and choose the groups from the pop up window
9. Tick or untick the Force Password change and Test Mode (Home Directory not working)
10. Click on Analyse (Prepare will not enable unless Analyse is opened), check then close Analyse window.
11. Click on Prepare
12. Ok the Prepare prompt, choose what columns to display if required
13. Look in the Preview window to see the extra columns created
14. Optional tick Auto-export passwords on completion
15. If happy then click on Import to import users
16. Export CSV to view data.

## Usage
Download or clone the repository and run the main PowerShell script:

```powershell
.\ADImportTool-BetaV1.ps1
