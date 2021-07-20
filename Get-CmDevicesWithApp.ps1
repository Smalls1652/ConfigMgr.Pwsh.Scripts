<#
.SYNOPSIS
    Search for devices from ConfigMgr with installed software.
.DESCRIPTION
    Search ConfigMgr for devices that have specific software installed. Data is pulled from the hardware inventory and includes the software name, the devices found, and the count of the devices.
.PARAMETER SiteCode
    The site code for the primary site.
.PARAMETER ProductName
    The name of the software to search for.
.PARAMETER SearchType
    The type of search to run. Valid options are: Explicit or Wildcard.
.EXAMPLE
    PS \> .\Get-CmDevicesWithApp.ps1 -SiteCode "ABC" -ProductName "Microsoft Project Professional 2016"

    Gets devices with software named 'Microsoft Project Professional 2016'.
.EXAMPLE
    PS \> .\Get-CmDevicesWithApp.ps1 -SiteCode "ABC" -ProductName "Microsoft Visio Professional" -SearchType "Wildcard"

    Gets devices with software named 'Microsoft Visio Professional' utilizing a wildcard search.
.OUTPUTS
    [CmSoftwareItem[]]
.NOTES
    --------

    - The 'ConfigurationManager' module needs to be imported into the current PowerShell session before running.

    - Wildcard searches do not need asterisks (*) in the 'ProductName' parameter. The wildcard search query is ran with a wildcard at the beginning and end of the 'ProductName' parameter.
        - For example if the 'ProductName' is set to "Visual Studio", then the wildcard search will be ran as "*Visual Studio*".
    
    --------
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0, Mandatory)]
    [string]$SiteCode,
    [Parameter(Position = 1, Mandatory)]
    [string]$ProductName,
    [Parameter(Position = 2)]
    [ValidateSet(
        "Explicit",
        "Wildcard"
    )]
    [string]$SearchType = "Explicit"
)

<#
    Class: VerboseProgressPrefs
    -------------------

    This class is just a simple way of storing Verbose/Progress preferences.
#>
class VerboseProgressPrefs {
    [System.Management.Automation.ActionPreference]$VerbosePref
    [System.Management.Automation.ActionPreference]$ProgressPref

    VerboseProgressPrefs([System.Management.Automation.ActionPreference]$verbosePref, [System.Management.Automation.ActionPreference]$progressPref) {
        $this.VerbosePref = $verbosePref
        $this.ProgressPref = $progressPref
    }
}

<#
    Class: CmSoftwareItem
    -------------------

    This class lists the software name, the devices with it, and the count of those devices.
#>
class CmSoftwareItem {
    [string]$SoftwareName
    [int]$DeviceCount
    [CmSoftwareItemDevice[]]$Devices

    CmSoftwareItem([string]$softwareName, [CmSoftwareItemDevice[]]$devices) {
        $this.SoftwareName = $softwareName
        $this.DeviceCount = ($devices | Measure-Object).Count
        $this.Devices = $devices
    }

    [string]ToString() {
        return "$($this.SoftwareName) / $($this.DeviceCount) devices"
    }
}

<#
    Class: CmSoftwareItemDevice
    -------------------

    This class is data about the device and the software it has.
#>
class CmSoftwareItemDevice {
    [string]$SoftwareName
    [string]$DeviceName

    [string]ToString() {
        return $this.DeviceName
    }
}

#Save the Verbose/Progress preferences before we continue.
#This is to stop the 'ConfigurationManager' module cmdlets from outputting Verbose/Progress lines, since supplying `-Verbose:$false` to those cmdlets doesn't work.
$savedVerbosePrefs = [VerboseProgressPrefs]::new($VerbosePreference, $ProgressPreference)

#Get all PSDrives with the PSProvider of 'CMSite' and then filter for the provided site code.
$cmSiteDrives = Get-PSDrive -PSProvider "CMSite" -ErrorAction "Stop"
$selectedSite = $cmSiteDrives | Where-Object { $PSItem.Name -eq $SiteCode }

#If '$selectedSite' is null, then throw and error.
if ($null -eq $selectedSite) {
    $PSCmdlet.ThrowTerminatingError(
        [System.Management.Automation.ErrorRecord]::new(
            [System.Exception]::new("Could not find '$($SiteCode)' as a PSDrive."),
            "SiteCodeNotFound",
            [System.Management.Automation.ErrorCategory]::ObjectNotFound,
            $cmSiteDrives
        )
    )
}

#Push the location to the site code's PSDrive. This is a requirement for 'ConfigurationManager' cmdlets to run.
Write-Verbose "Pushing location to '$($SiteCode):'."
Push-Location -Path "$($SiteCode):" -StackName "CmScript" -Verbose:$false

#This `try/finally` statement is to handle cancellations/errors, so that the original path location is returned back. 
try {
    #Generate two WQL queries to get all of the software from the software inventory that matches the product name.
    $productQueries = $null
    switch ($SearchType) {
        "Wildcard" {
            Write-Verbose "Using wildcard search."
            $productQueries = @(
                "SELECT * FROM SMS_G_System_Add_Remove_Programs WHERE DisplayName LIKE '%$($ProductName)%'",
                "SELECT * FROM SMS_G_System_Add_Remove_Programs_64 WHERE DisplayName LIKE '%$($ProductName)%'"
            )
            break
        }

        Default {
            Write-Verbose "Using explicit search."
            $productQueries = @(
                "SELECT * FROM SMS_G_System_Add_Remove_Programs WHERE DisplayName = '$($ProductName)'",
                "SELECT * FROM SMS_G_System_Add_Remove_Programs_64 WHERE DisplayName = '$($ProductName)'"
            )
            break
        }
    }

    #Run the queries to get the data.
    Write-Verbose "Running queries against 'Add/Remove Programs' hardware classes."
    $softwareFound = foreach ($query in $productQueries) {
        #Disable Verbose/Progress output temporarily.
        $VerbosePreference = "SilentlyContinue"
        $ProgressPreference = "SilentlyContinue"

        $foundResults = Invoke-CMWmiQuery -Query $query

        #Restore Verbose/Progress output preferences.
        $VerbosePreference = $savedVerbosePrefs.VerbosePref
        $ProgressPreference = $savedVerbosePrefs.ProgressPref

        foreach ($item in $foundResults) {
            $item
        }
    }

    #Create a WQL query string and write it to the verbose output. This query is specialized for creating a device collection based off of it.
    #$WqlQueryForCollection = "Select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System INNER JOIN SMS_G_System_Add_Remove_Programs ON SMS_G_System_Add_Remove_Programs.ResourceID = SMS_R_System.ResourceID INNER JOIN SMS_G_System_Add_Remove_Programs_64 ON SMS_G_System_Add_Remove_Programs_64.ResourceID = SMS_R_System.ResourceID where SMS_G_System_Add_Remove_Programs.DisplayName = '$($ProductName)' OR SMS_G_System_Add_Remove_Programs_64.DisplayName = '$($ProductName)'"

    #Get the unique software names and the unique resource IDs for each system.
    $uniqueSoftwareNames = ($softwareFound | Select-Object -Property "DisplayName" -Unique).DisplayName | Sort-Object
    $softwareFound = $softwareFound | Sort-Object -Property "ResourceID" -Unique

    foreach ($software in $uniqueSoftwareNames) {
        Write-Verbose "Getting devices with '$($software)'."

        #Get the device names that have the software.
        $devicesWithSoftware = foreach ($item in ($softwareFound | Where-Object { $PSItem.DisplayName -eq $software })) {
            #Disable Verbose/Progress output temporarily.
            $VerbosePreference = "SilentlyContinue"
            $ProgressPreference = "SilentlyContinue"

            #We could run `Get-CmDevice`; however, it takes a bit more time to run that since it returns more data than needed. Even if you use `-Fast` it takes too long to run.
            #Instead we're going to run a custom WQL query to only return the name of the system.
            $deviceName = (Invoke-CMWmiQuery -Query "Select Name from SMS_R_System WHERE ResourceId = $($item.ResourceID) and ResourceType = 5").Name

            #Restore Verbose/Progress output preferences.
            $VerbosePreference = $savedVerbosePrefs.VerbosePref
            $ProgressPreference = $savedVerbosePrefs.ProgressPref

            [CmSoftwareItemDevice]@{
                "SoftwareName" = $software;
                "DeviceName"   = $deviceName;
            }
        }

        #Create the 'CmSoftwareItem' object and write to the output.
        $softwareItemObj = [CmSoftwareItem]::new($software, $devicesWithSoftware)
        Write-Output -InputObject $softwareItemObj
    }
}
finally {
    #Return the location back to the original path before script execution.
    Pop-Location -StackName "CmScript"
}