# AX2009PS
Dynamics AX 2009 PowerShell Scripts

Pre-requisite
----------------
- Put the Microsoft.Dynamics.BusinessConnectorNet.dll in the same folder where you store the module
- Put a valid .axc file in the same folder where you store the module
- Must run with 32bit version of Powershell

Summary
----------------
The module is written under the assumption that the businessConnector dll and a valid .axc to be available with the script module.
The module will copy those 2 files to C:\Temp if they are not already available. Feel free to change the code in Get-AXObject function to suits your need.

Once the connection to AX is settled. Play around with the New-AXSelectStmt function. 

```Powershell
Get-help New-AXSelectStmt 
```

to see more information.

Available Functions
----------------
- Use-Culture 
- Convert-Filename 
- Get-AXObject 
- Get-AXInfo 
- New-AXSelectStmt 

Installation
----------------
Put the .psd1 and .psm1 into a folder named "AX2009PS". Put also the businessConnector dll and a valid .axc in the same folder.
Then copy to folder to one of the path listed under $env:PSModulePath.


