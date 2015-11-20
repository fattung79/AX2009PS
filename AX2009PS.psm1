$Global:logOffAXFailed = "-- Failed to log off AX Session --"
$Global:defaultCompany = "ceu"

Function Use-Culture
{
    #############################################################################
    ##
    ## Use-Culture
    ##
    ## From Windows PowerShell Cookbook (O'Reilly)
    ## by Lee Holmes (http://www.leeholmes.com/guide)
    ##
    #############################################################################

    <#

    .SYNOPSIS

    Invoke a scriptblock under the given culture

    .EXAMPLE

    Use-Culture fr-FR { [DateTime]::Parse("25/12/2007") }
    mardi 25 decembre 2007 00:00:00

    #>

    param(
        ## The culture in which to evaluate the given script block
        [Parameter(Mandatory = $true)]
        [System.Globalization.CultureInfo] $Culture,

        ## The code to invoke in the context of the given culture
        [Parameter(Mandatory = $true)]
        [ScriptBlock] $ScriptBlock
    )

    Set-StrictMode -Version Latest

    ## A helper function to set the current culture
    function Set-Culture([System.Globalization.CultureInfo] $culture)
    {
        [System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture
        [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
    }

    ## Remember the original culture information
    $oldCulture = [System.Threading.Thread]::CurrentThread.CurrentUICulture

    ## Restore the original culture information if
    ## the user's script encounters errors.
    trap { Set-Culture $oldCulture }

    ## Set the current culture to the user's provided
    ## culture.
    Set-Culture $culture

    ## Invoke the user's scriptblock
    & $ScriptBlock

    ## Restore the original culture information.
    Set-Culture $oldCulture
}

Function Convert-FileName
{
    <#

    .SYNOPSIS

    Add date time suffix to filename

    .EXAMPLE

    Convert-Filename "proceess.log"
    
        \process_20151025_0423.log
    #>
    [cmdletBinding()]
    param
    (
        [parameter(Mandatory=$true)] [string] $fileLoc
    )

    $filePath = Split-Path $fileLoc
    $dateStr = Get-Date -Format "yyyyMMdd_hhmm"
    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($fileLoc) + "_" + $dateStr
    $fileExt = [System.IO.Path]::GetExtension($fileLoc)
    $fileLoc = $filePath + "\" + $fileName + $fileExt

    $fileLoc
}

Function Get-AXObject
{
    <#

    .SYNOPSIS

    Create a Microsoft.Dynamics.BusinessConnector.Axapta object, login to AX and returns the object. A client configuration file (.axc) named AX2009.axc should be put available the module's folder.
    The Microsoft.Dynamics.BusinessConnectorNet.dll should also be put in the folder 

    .EXAMPLE

    $ax = Get-AXObject | Get-Member    
    
       TypeName: Microsoft.Dynamics.BusinessConnectorNet.Axapta

    Name                         MemberType Definition                                                                                                                                                                     
    ----                         ---------- ----------                                                                                                                                                                     
    CallJob                      Method     System.Void CallJob(string jobName, Microsoft.Dynamics.BusinessConnectorNet.AxaptaObject argsObject), System.Void CallJob(string jobName)                                      
    CallStaticClassMethod        Method     System.Object CallStaticClassMethod(string className, string methodName, System.Object param1, System.Object param2, System.Object param3), System.Object CallStaticClassMet...
    CallStaticRecordMethod       Method     System.Object CallStaticRecordMethod(string recordName, string methodName, System.Object param1, System.Object param2, System.Object param3), System.Object CallStaticRecord...
    CreateAxaptaBuffer           Method     Microsoft.Dynamics.BusinessConnectorNet.AxaptaBuffer CreateAxaptaBuffer()                                                                                                      
    CreateAxaptaContainer        Method     Microsoft.Dynamics.BusinessConnectorNet.AxaptaContainer CreateAxaptaContainer()                                                                                                
    CreateAxaptaObject           Method     Microsoft.Dynamics.BusinessConnectorNet.AxaptaObject CreateAxaptaObject(string className, System.Object param1, System.Object param2, System.Object param3), Microsoft.Dynam...
    CreateAxaptaRecord           Method     Microsoft.Dynamics.BusinessConnectorNet.AxaptaRecord CreateAxaptaRecord(string recordName), Microsoft.Dynamics.BusinessConnectorNet.AxaptaRecord CreateAxaptaRecord(int reco...
    Dispose                      Method     System.Void Dispose()                                                                                                                                                          
    Equals                       Method     bool Equals(System.Object obj)                                                                                                                                                 
    ExecuteStmt                  Method     System.Void ExecuteStmt(string statement, Microsoft.Dynamics.BusinessConnectorNet.AxaptaRecord param1, Microsoft.Dynamics.BusinessConnectorNet.AxaptaRecord param2, Microsof...
    GetBufferCount               Method     int GetBufferCount()                                                                                                                                                           
    GetContainerCount            Method     int GetContainerCount()                                                                                                                                                        
    GetHashCode                  Method     int GetHashCode()                                                                                                                                                              
    GetLoggedOnAxaptaObjectCount Method     int GetLoggedOnAxaptaObjectCount()                                                                                                                                             
    GetObject                    Method     Microsoft.Dynamics.BusinessConnectorNet.AxaptaObject GetObject(string objectName)                                                                                              
    GetObjectCount               Method     int GetObjectCount()                                                                                                                                                           
    GetRecordCount               Method     int GetRecordCount()                                                                                                                                                           
    GetType                      Method     type GetType()                                                                                                                                                                 
    Logoff                       Method     bool Logoff()                                                                                                                                                                  
    Logon                        Method     System.Void Logon(string company, string language, string objectServer, string configuration)                                                                                  
    LogonAs                      Method     System.Void LogonAs(string user, string domain, System.Net.NetworkCredential bcProxyCredentials, string company, string language, string objectServer, string configuration)   
    LogonAsGuest                 Method     System.Void LogonAsGuest(System.Net.NetworkCredential bcProxyCredentials, string company, string language, string objectServer, string configuration)                          
    Refresh                      Method     System.Void Refresh()                                                                                                                                                          
    Session                      Method     int Session()                                                                                                                                                                  
    ToString                     Method     string ToString()                                                                                                                                                              
    TTSAbort                     Method     System.Void TTSAbort()                                                                                                                                                         
    TTSBegin                     Method     System.Void TTSBegin()                                                                                                                                                         
    TTSCommit                    Method     System.Void TTSCommit()                                                                                                                                                        
    HttpContextAccessible        Property   System.Boolean HttpContextAccessible {get;}    
    #>
    
    [cmdletBinding()]
    param
    (        
        [parameter(Mandatory=$false)] [string] $company = $defaultCompany,
        [parameter(Mandatory=$false)] [string] $language = "",
        [parameter(Mandatory=$false)] [string] $aos = "",
        [parameter(Mandatory=$false)] [string] $config = ""      
    )
    
    #region Setup DLL
    $tempPath = "C:\temp"

    $wmi = Get-WmiObject -Class win32_computersystem -ComputerName localhost
    $hostname = $wmi.DNSHostName
    $axcName = "\AX2009.axc"
    $dllName = "\Microsoft.Dynamics.BusinessConnectorNet.dll"
    $modPath = (Get-Module AX2009PS).ModuleBase
    $fullPath = $modPath + $dllName    
    $targetPath = $tempPath + $dllName

    if (-Not (Test-Path($targetPath)))
    {
        Copy-Item $fullPath $tempPath | Out-Null
    }    

    $axcPath = $modPath + $axcName
    $axcTargetPath = $tempPath + $axcName
    if (-Not (Test-Path($axcTargetPath)))
    {
        Copy-Item $axcPath $tempPath | Out-Null
    }

    [reflection.Assembly]::Loadfile($targetPath) | Out-Null
    #endregion    

    $ax = new-object Microsoft.Dynamics.BusinessConnectorNet.Axapta
    if ($config -eq "")
    {
        $config = $axcTargetPath
    }
    $ax.logon($company,$language,$aos,$config)

    return $ax
}

Function New-AXSelectStmt
{
    <#

    .SYNOPSIS

    Allow user to run query against AX by passing in table name(s) and the select statement (X++ syntax). By default this query return an array of PSObjects. Each object contains the first 10 fields (per table) it can find.
    It is possible to specify which fields to include in the PSObject by passing in a comma seperated string to FieldList parameter (per table).    

    .EXAMPLE

    New-AXSelectStmt CustTable -stmt "SELECT * FROM %1"
    
    Customer account : 9100
    Name             : Mike Miller
    Address          : 
    Telephone        : 
    Fax              : 
    Invoice account  : 
    Customer group   : 90
    Line discount    : 
    Terms of payment : N007
    Cash discount    : 

    Customer account : Contoso
    Name             : Contoso Standard Template
    Address          : 
    Telephone        : 
    Fax              : 
    Invoice account  : 
    Customer group   : 80
    Line discount    : 
    Terms of payment : N010
    Cash discount    : 
    
    ...
    
    .EXAMPLE
    
    New-AXSelectStmt CustTable -company "ceu" -stmt "SELECT * FROM %1 WHERE %1.CustGroup == '10'"
    
    Customer account : 1304
    Name             : Otter Wholesales
    Address          : 123 Peach Road Federal Way, WA 98003 US
    Telephone        : 123-555-0170
    Fax              : 321-555-0159
    Invoice account  : 
    Customer group   : 10
    Line discount    : 
    Terms of payment : P007
    Cash discount    : 

    Customer account : 9024
    Name             : Dolphin Wholesales
    Address          : 
    Telephone        : 111-555-0114
    Fax              : 111-555-0115
    Invoice account  : 
    Customer group   : 10
    Line discount    : 
    Terms of payment : N060
    Cash discount    : 
    
    ...
    
    .EXAMPLE
    
    New-AXSelectStmt SalesTable,SalesLine -company "ceu" -stmt "SELECT * FROM %1 JOIN %2 WHERE %1.SalesId == %2.SalesId" -top 10 -fieldlists "SalesId,SalesName,CustAccount","ItemId,SalesPrice,SalesQty,LineAmount"
    
    SalesTable_Sales order      : SO-100005
    SalesTable_Name             : Contoso Retail Seattle
    SalesTable_Customer account : 3002
    SalesLine_Item number       : 1151
    SalesLine_Unit price        : 62.21
    SalesLine_Quantity          : 12
    SalesLine_Net amount        : 746.52

    SalesTable_Sales order      : SO-100005
    SalesTable_Name             : Contoso Retail Seattle
    SalesTable_Customer account : 3002
    SalesLine_Item number       : 1153
    SalesLine_Unit price        : 61.81
    SalesLine_Quantity          : 2
    SalesLine_Net amount        : 123.62
    
    ...
    #>
    [cmdletBinding()]
    param
    (
        [parameter(Mandatory=$true)] 
        [string[]] $tables,
        [string] $stmt,
        
        [parameter(Mandatory=$false)] 
        [string[]] $fieldLists,
        [string] $separator = ",",
        [string] $top = 0,
        [string] $numOfFields = 10,        
        [string] $company = "",
        [string] $language = "",
        [string] $aos = "",
        [string] $config = "",
        [switch] $showLabel
    )

    Try
    {
        $ax = Get-AXObject -company $company -language $language -aos $aos -config $config
        
        $tableBuffers = {@()}.Invoke()
        #$tableList = $tables.split(",")
        foreach ($t in $tables)
        {
            $buffer = $ax.CreateAxaptaRecord($t)
            $tableBuffers.Add($buffer)
        }

        $ax.ExecuteStmt($stmt,$tableBuffers)
            
        $list = {@()}.Invoke()      
        $recCount = 0

        Do { 
            $obj = New-Object PSObject; 
            if ($fieldlists.Count -eq 0 -or $fieldLists -eq $null)
            {
                for ([int] $j=0; $j -lt $tablebuffers.Count; $j++)
                {
                    $i = 1
                    $fieldCounts = 0
                    Do
                    {
                        if ($tableBuffers[$j].get_fieldLabel($i) -eq "UNKNOWN")
                        {
                            $i++
                            continue;
                        }

                        if ($showLabel)
                        {
                            $fieldLabel = $tableBuffers[$j].FieldLabel($i)
                            if ($tables.Count -gt 1)
                            {
                                $fieldLabel = $tables[$j] + "_" + $fieldLabel
                            }
                        }
                        else
                        {                                                        
                            $dictField = $ax.CreateAxaptaObject("SysDictField",$tableBuffers[$j].field("tableId"),$i)
                            $fieldLabel = $dictField.Call("name")
                            if ($tables.Count -gt 1)
                            {
                                $fieldLabel = $tables[$j] + "_" + $fieldLabel
                            }
                        }
                                                                        
                        $obj | Add-Member -Name $fieldLabel -Value $tableBuffers[$j].Field($i) -MemberType NoteProperty; 
                        $fieldCounts++
                        $i++
                    } while ($fieldCounts -lt $numOfFields)                                           
                } 
            }
            else
            {
                for ([int] $j=0; $j -lt $tablebuffers.Count; $j++)
                {
                    $fields = $fieldLists[$j].Split($separator)
                    Foreach ($f in $fields)
                    {
                        if ($tableBuffers[$j].get_fieldLabel($f) -eq "UNKNOWN")
                        {
                            continue;
                        }

                        if ($showLabel)
                        {
                            $fieldLabel = $tableBuffers[$j].FieldLabel($f)
                            if ($tables.Count -gt 1)
                            {
                                $fieldLabel = $tables[$j] + "_" + $fieldLabel
                            }
                        }
                        else
                        {
                            $fieldLabel = $f
                            if ($tables.Count -gt 1)
                            {
                                $fieldLabel = $tables[$j] + "_" + $fieldLabel
                            }                        
                        }
                        $obj | Add-Member -Name $fieldLabel -Value $tableBuffers[$j].Field($f) -MemberType NoteProperty;
                    }
                }
            }
            $list.Add($obj) 
            $recCount++
            if (($top -ne 0) -AND ($recCount -eq $top))
            {
                break
            }
        } while ($tableBuffers[0].Next())
        

        $list
    }
    Catch [System.Exception]
    {
        "Caught an Exception"
        $_.Exception|format-list -force
    } 
    Finally 
    {
        if(-not $ax.Logoff())        
        {
            $Global:logOffAXFailed
        }
    }    
}

Function Get-AXInfo
{    
    <#

    .SYNOPSIS

    Get information about the AOS connected to

    .EXAMPLE

    Get-AXInfo
    
    AOS     : AX5-W8-01
    Build No: 1500.2985
    Instance: 01
    Port    : 2712
    
    #>
    [cmdletBinding()]
    param
    (        
        [parameter(Mandatory=$false)] [string] $company = "",
        [parameter(Mandatory=$false)] [string] $language = "",
        [parameter(Mandatory=$false)] [string] $aos = "",
        [parameter(Mandatory=$false)] [string] $config = ""        
    )

    Try {    
        $ax = Get-AXObject -company $company -language $language -aos $aos -config $config

        $xSession = $ax.CreateAxaptaObject("XSession")
        "AOS     : " + $xSession.call("AOSName")     
    
        $xApplication = $ax.CreateAxaptaObject("XApplication")   
        "Build No: " + $xApplication.call("buildNo")           
        "Instance: " + $ax.CallStaticClassMethod("Session", "getAOSInstance")     
        "Port    : " + $ax.CallStaticClassMethod("Session", "getAOSPort") 
    } 
    Catch [System.Exception]
    {
        "Caught an Exception"
        $_.Exception|format-list -force
    } 
    Finally 
    {
        if(-not $ax.Logoff())
        {
            $Global:logOffAXFailed
        }
    }
}
