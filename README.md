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

Available Functions
----------------
Use-Culture - a helper function written by Lee Holmes to run a scriptblock under a given culture
Convert-Filename - a helper function used to append dateTime suffix to a filename
Get-AXObject - a helper function that returns a valid Microsoft.Dynamics.BusinessConnector.Axapta object, which has been logged in
Get-AXInfo - Connect to an AOS and returns it's information
New-AXSelectStmt - Allow user to run query against AX by passing in table name(s) and the select statement (X++ syntax). By default this query return an array of PSObjects.

Installation
----------------
Put the .psd1 and .psm1 into a folder named "AX2009PS". Put also the businessConnector dll and a valid .axc in the same folder.
Then copy to folder to one of the path listed under $env:PSModulePath.

Sample code
----------------
Import-Module AX2009PS
New-AXSelectStmt SalesTable,SalesLine -company "ceu" -stmt "SELECT * FROM %1 JOIN %2 WHERE %1.SalesId == %2.SalesId" -top 10 -fieldlists "SalesId,SalesName,CustAccount","ItemId,SalesPrice,SalesQty,LineAmount"

