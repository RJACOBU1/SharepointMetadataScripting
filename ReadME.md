# SharePoint Metadata Scripting

## Overview
This project aims to automate the process of adding metadata to SharePoint files using PowerShell scripts. The metadata added to these files serves as the backbone for the MAT tool, which utilizes this information to facilitate easy file retrieval and management.

## Goals
- Automate metadata addition to SharePoint files.
- Enhance file discoverability through structured metadata.
- Support the MAT tool in efficiently locating files based on metadata.

## Usage
To use the scripts in this repository, ensure you have the necessary permissions and PowerShell installed.

### Installing PowerShell
To install or update to the latest version of PowerShell, you can use the Windows Package Manager (winget). Run the following commands in your command line interface:

```powershell
winget search Microsoft.PowerShell
winget install --id Microsoft.PowerShell --source winget
```

### Installing PnP.PowerShell Module
The PnP.PowerShell module is required to run the scripts in this repository. Install it by executing the following command in your PowerShell:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```


