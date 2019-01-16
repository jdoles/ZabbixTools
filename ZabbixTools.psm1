<#
    ZabbixTools.psm1
    Updated: 2018-03-08
    Author: Justin Doles
    Requires: PowerShell 5.1 or higher, Internet Explorer 11 or higher
    Description: Library of tools to work with Zabbix API https://www.zabbix.com/documentation/3.2/manual/api
    Notes:
    MUST execute IE once as the user who runs this script or you'll receive the error below
    System.NotSupportedException: The response content cannot be parsed because the Internet Explorer engine is not available, 
    or Internet Explorer's first-launch configuration is not complete. Specify the UseBasicParsing parameter and try again.
#>
<#
.SYNOPSIS
Loads the configuration File

.DESCRIPTION
Loads a JSON formatted configuration file from the current directory

.PARAMETER ConfigFile
JSON formatted configuration file

.EXAMPLE
LoadConfig "ZabbixTools.json"

.NOTES
None
#>
function LoadConfig {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$ConfigFile
    )

    # Get the local path
    $configpath = $PSScriptRoot + "\" + $ConfigFile
    if (Test-Path $configpath) {
        return Get-Content $configpath | ConvertFrom-Json
    } else {
        return $false
    }
}

<#
.SYNOPSIS
Posts data to a URI

.DESCRIPTION
Posts JSON data to the specified URI and returns a Powershell object if successful

.PARAMETER URI
URI to post to to

.PARAMETER Body
Body of request

.EXAMPLE
An example

.NOTES
General notes
#>
function PostJSONData {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$URI,
        [Parameter(Mandatory=$true)]
        [string]$Body
    )

    try {
        [Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"
        # Write-Host "Connecting to:"$URI
        $response = Invoke-WebRequest -Uri $URI -Method Post -Body $Body -ContentType "application/json"
        # Write-Host "Status: "$response.StatusCode
        if ($response.StatusCode -eq 200) {
            return $response.Content | ConvertFrom-Json
        } else {
            return $false
        }
    } catch {
        $e = $_.Exception.GetType().Name
        Write-Host $_.Exception
        Write-Host "PostJSONData threw"$e
        return $false
    }    
}

<#
.SYNOPSIS
Gets an API key

.DESCRIPTION
Gets an API key for the the user specifed

.PARAMETER URI
Zabbix server URI

.PARAMETER Username
Zabbix user name

.PARAMETER Password
Zabbix user password

.EXAMPLE
An example

.NOTES
General notes
#>
function Get-APIKey {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$URI,
        [Parameter(Mandatory=$true)]
        [string]$Username,
        [Parameter(Mandatory=$true)]
        [string]$Password
    )
    $body = @{
        jsonrpc = "2.0"
        method = "user.login"
        id = 1
        auth = $null
        params = @{
            user = $Username
            password = $Password
        }
    } | ConvertTo-Json
    
    if ($result = PostJSONData $URI $body) {
        if ($result.result.Length -gt 0) {
            # Got a result
            return $result.result
        } else {
            # No results
            return $false
        }
    } else {
        # Call may have failed
        return $false
    }
}

<#
.SYNOPSIS
Gets the specified Windows host from Zabbix

.DESCRIPTION
Fetches the specified host from Zabbix and it's associated interfaces.  This requires the inventory OS field to be populated with Windows.

.PARAMETER HostName
Name of the host to search for

.PARAMETER URI
Zabbix server URI

.PARAMETER APIKey
API key for Zabbix

.EXAMPLE
Nah

.NOTES
General notes
#>
function Get-WindowsHostFromZabbix {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$HostName,
        [Parameter(Mandatory=$true)]
        [string]$URI,
        [Parameter(Mandatory=$true)]
        [string]$APIKey
    )
    $body = @{
        jsonrpc = "2.0"
        method = "host.get"
        auth = $APIKey
        id = 1
        params = @{
            output = @("host", "status")
            selectInterfaces = "extend"
            selectInventory = @("os")
            searchInventory = @{
                os = "Windows"
            }
            filter = @{
                host = @($HostName)
            }
        }
    } | ConvertTo-Json -Depth 4
    
    if ($result = PostJSONData $URI $body) {
        if ($result.result.Length -gt 0) {
            # Got a result
            return $result.result[0]
        } else {
            # No results
            return $false
        }
    } else {
        # Call may have failed
        return $false
    }
}

<#
.SYNOPSIS
Gets a list of computers from Active Directory

.DESCRIPTION
Gets a list of computers from the specified OU

.PARAMETER SearchBase
LDAP path to search. Example: ou=servers,dc=example,dc=com

.PARAMETER DomainController
FQDN of the domain controller. This is optional.

.EXAMPLE
An example

.NOTES
General notes
#>
function Get-ComputersFromAD {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$SearchBase,
        [Parameter(Mandatory=$false)]
        [string]$DomainController = $null
    )

    try {
        # Grab a list of enabled computers from the specified search base
        $computers = $null
        if ($DomainController) {
            $computers = Get-ADComputer -LDAPFilter "(&(objectClass=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))" -SearchBase $SearchBase -Server $DomainController
        } else {
            $computers = Get-ADComputer -LDAPFilter "(&(objectClass=computer)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))" -SearchBase $SearchBase
        }        
        return $computers
    } catch {
        # Something went wrong
        $e = $_.Exception.GetType().Name
        Write-Host "getComputersFromAD:" $e ":" $_.Exception.Message
        return $false
    }
}

<#
.SYNOPSIS
Creates a new host

.DESCRIPTION
Creates a new host in Zabbix with the specified parameters

.PARAMETER HostName
Parameter description

.PARAMETER UseIP
Parameter description

.PARAMETER IP
Parameter description

.PARAMETER DNS
Parameter description

.PARAMETER Port
Parameter description

.PARAMETER GroupID
Parameter description

.PARAMETER InventoryMode
Parameter description

.PARAMETER URI
Parameter description

.PARAMETER APIKey
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
function New-ZabbixHost {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [string]$HostName,        
        
        [Parameter(Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [int]$UseIP,
        
        [Parameter(Mandatory=$false,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [string]$IP = "127.0.0.1",
        
        [Parameter(Mandatory=$false,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [string]$DNS = "",
        
        [Parameter(Mandatory=$false,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [string]$Port = 10050,
        
        [Parameter(Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [string]$GroupID,

        [Parameter(Mandatory=$false,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [string]$InventoryMode = 0,

        [Parameter(Mandatory=$true)]
        [string]$URI,

        [Parameter(Mandatory=$true)]
        [string]$APIKey
    )

    Process {
        # Prepare the request object and convert to JSON
        $body = @{
            jsonrpc = "2.0"
            method = "host.create"
            auth = $APIKey
            id = 1
            params = @{
                host = $HostName
                interfaces = @(
                    @{
                        type = 1
                        main = 1
                        useip = $UseIP
                        ip = $IP
                        dns = $DNS
                        port = $Port
                    }
                )
                groups = @(
                    @{
                        groupid = $GroupID
                    }
                )
                templates = @()
                inventory_mode = $InventoryMode
                inventory = @{}
            }
        } | ConvertTo-Json -Depth 100
        
        if ($result = PostJSONData $URI $body) {
            if ($result.result) {
                # Got a result
                Write-Verbose ("Host " + $HostName + " with ID " + $result.result.hostids[0] + " created")
                return $result.result.hostids[0]
            } elseif ($result.error) {
                # Possible error
                Write-Error ("An error occurred. " + $result.error.data)
                Write-Verbose ("Error " + $result.error.code.ToString() + ". " + $result.error.message + " " + $result.error.data)
                return $false
            }
        } else {
            # Call may have failed
            Write-Verbose "Call failed"
            return $false
        }
    }
}

<#
.SYNOPSIS
Compares computers in AD to those in Zabbix

.DESCRIPTION
Takes a list of computers from a specified OU and compares them to the hosts in Zabbix

.PARAMETER URI
Zabbix server URI

.PARAMETER APIKey
API key for Zabbix

.PARAMETER SearchBase
LDAP path to search. Example: ou=servers,dc=example,dc=com

.PARAMETER DomainController
FQDN of the domain controller. This is optional.

.PARAMETER ExcludedComputers
A regular expression containing a list of servers to skip.  Example: server*|workstation*|bob1

.EXAMPLE
An example

.NOTES
General notes
#>
function CompareADToZabbix {
    [CmdletBinding()]
    [OutputType([System.Collections.ArrayList])]
    Param (
        [Parameter(Mandatory=$true)]
        [string]$URI,
        [Parameter(Mandatory=$true)]
        [string]$APIKey,
        [Parameter(Mandatory=$true)]
        [string]$SearchBase,
        [Parameter(Mandatory=$false)]
        [string]$DomainController = $null,
        [Parameter(Mandatory=$false)]
        [string]$ExcludedComputers = $null
    )
    
    # Get the list of AD computers
    $adcomputers = Get-ComputersFromAD $SearchBase $DomainController
    # Holds the list of missing computers
    $missingcomputers = New-Object System.Collections.ArrayList
    foreach ($computer in $adcomputers) {
        if ($computer.Name.ToLower() -notmatch $ExcludedComputers) {
            if (Get-WindowsHostFromZabbix $computer.Name.ToLower() $URI $APIKey) { 
                # Do nothing
            } else {
                # Computer missing from Zabbix
                [void]$missingcomputers.Add($computer)
            }
        }
    }

    return $missingcomputers
}
