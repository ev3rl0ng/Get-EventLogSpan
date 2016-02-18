Function Get-EventLogSpan {

<#
    .SYNOPSIS
    Function for veryfing Windows events log span - how old is the oldest available log event entry

    .DESCRIPTION
    Function for veryfing Windows events log span - how old is the oldest available log event entry
    As a result for this function you receive range between the oldest event log entry and the newest one.

    .PARAMETER ComputerName
    Computer on which logs need to queried.
    By default localhost will be queried.

    .PARAMETER LogsScope
    You can select between Classic and All - default is Classic,
    For basic installation Windows Server 2008 R2 these logs are: Application, Security, Setup, System, Forwarded Events.
    All logs means that also "Applications and Services Logs" will be included - query all logs can be time consuming because LogParse can't be used for it.

    .PARAMETER ExcludeEmptyLogs
    By default empty logs are excluded from results.

    .PARAMETER WarningLevelDays
    The amount od days for which log can be marked in results as Warning if the span of logs is smaller than warning level.
    Default warning level is set to 30 days.

    .PARAMETER CriticalLevelDays
    The amount od days for which log can be marked in results as Critical if the span of logs is smaller than critical level.
    Default critical level is set to 7 days.

    .PARAMETER LogParserInstalled
    Use this parameter (set to $false) if Log Parser is not installed, the native PowerShell cmdlet will be used to parse all event logs.
    If set to $false log parsing will be much slower - due to use PowerShell native command to find the oldes log event entry.

    .EXAMPLE

    Get events log span from the localhost

    [PS] > Get-Date
    Wednesday, February 17, 2016 11:26:20 PM

    [PS] > Get-EventLogSpan -ComputerName localhost -LogsScope Classic -WarningLevelDays 14

    LogName         : Application
    LogSpanStatus   : Warning
    ComputerName    : localhost
    OldestEventTime : 2016-02-10 00:45:13
    LogTimeSpan     : 7.22:41:7.7739014

    LogName         : Security
    LogSpanStatus   : Normal
    ComputerName    : localhost
    OldestEventTime : 2015-12-06 00:44:55
    LogTimeSpan     : 73.22:41:25.5087764

    LogName         : System
    LogSpanStatus   : Critical
    ComputerName    : localhost
    OldestEventTime : 2016-02-14 13:11:55
    LogTimeSpan     : 3.10:14:25.1181514

    .LINK
    https://github.com/it-praktyk/Get-EventLogSpan

    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech

    .NOTES

    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net

    CONTRIBUTORS:
    - Thomas Rhoads, ev3rl0ng[at]gmail[dot]com

    KEYWORDS: Windows, PowerShell, EventLogs

    BASE REPOSITORY: https://github.com/it-praktyk/Get-EventLogSpan

    TODO
    - add pipeline support for ComputerName
    - update help INPUT/OUTPUT description


    VERSION HISTORY
    - 0.3.0 - 2015-01-20 - first version published on GitHub
    - 0.4.0 - 2015-01-21 - output updated - now include timespan, warning and critical levels added as parameters, output can be coloured
    - 0.5.0 - 2015-01-25 - checking not classic logs corrected,progress indicator added, lots improvments added
    - 0.5.1 - 2015-01-26 - checking oldest log corrected for remote computers
    - 0.5.2 - 2015-01-26 - checking using Log Parser corrected, output for status for empty logs corrected
    - 0.5.3 - 2015-02-05 - double quoute to single quote changed for static strings, minor updates
    - 0.6.0 - 2015-02-06 - check if .Net Framework 3.5 is installed - needed for Get-WinEvent cmdlet
    - 0.6.1 - 2015-02-09 - help updated
    - 0.6.2 - 2015-02-09 - script updated due to warning displayed by Script Analyzer e.g. positional parameter changed to named etc.
    - 0.7.0 - 2015-02-10 - minor bugs corrected, tabs replaced to 4 spaces to normalize looks between editors, first version published on TechNet
    - 0.8.0 - 2015-02-10 - query used for query data from remote computers by logparser corrected
    - 0.9.0 - 2016-02-17 - Thomas Rhoads (ev3rl0ng[at]gmail[dot]com) - added check for administrator token.
    - 0.9.1 - 2016-02-17 - Thomas Rhoads (ev3rl0ng[at]gmail[dot]com) - moved .net check to Begin section and added support for .Net > 3.5
    - 0.9.2 - 2016-02-17 - Thomas Rhoads (ev3rl0ng[at]gmail[dot]com) - Removed Test-Key function as it is no longer used.
    - 1.0.0 - 2016-02-17 - The license changed to MIT, the parameter OutputDirection removed, the function reformatted, by default output is returned as PowerShell object
    - 1.0.1 - 2016-02-18 - Thomas Rhoads (ev3rl0ng[at]gmail[dot]com) - Added - to Property in Select-Object portion of .NET 3.5 Check. Fixed Bug.
    - 1.1.0 - 2016-02-18 - Thomas Rhoads (ev3rl0ng[at]gmail[dot]com) - Added basic ping test to prevent error deluge when remote computer is unreachable.

    LICENSE
    Copyright (c) 2016 Wojciech Sciesinski
    This function is licensed under The MIT License (MIT)
    Full license text: https://opensource.org/licenses/MIT

   #>


#Requires -Version 2.0

[CmdletBinding()]

param (
    [parameter(mandatory=$false,Position=0)]
    [String]$ComputerName='localhost',

    [parameter(mandatory=$false,Position=1)]
    [ValidateSet('All','Classic')]
    [String]$LogsScope='Classic',

    [parameter(mandatory=$false,Position=2)]
    [Bool]$ExcludeEmptyLogs=$true,

    [parameter(mandatory=$false,Position=3)]
    [Int32]$WarningLevelDays=30,

    [parameter(mandatory=$false,Position=4)]
    [Int32]$CriticalLevelDays=7,

    [parameter(mandatory=$false,Position=5)]
    [Bool]$ColourOutput=$false,

    [parameter(mandatory=$false,Position=6)]
    [Bool]$LogParserInstalled=$true

)

Begin {

        #Set-StrictMode -Version 2

        # Get currently logged on principal.
        $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

        # Check for Administrator Rights bit.
        If (!$currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )) {

            Write-Error -Message "This function requires elevation. Please run PowerShell as an Administrator"

            Break
        }

        $currentPrincipal = $null

        # Iterate through all .NET versions installed and determine whether we have a satisfactory version.
        # Based on answer at: http://stackoverflow.com/questions/3487265/powershell-script-to-return-versions-of-net-framework-on-a-machine
        $boolNetFXVersionOK = $false

        Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -recurse |
        Get-ItemProperty -name Version,Release -ErrorAction SilentlyContinue |
        Where-Object -FilterScript { $_.PSChildName -match '^(?!S)\p{L}'} |
        Select-Object -Property PSChildName, Version, Release |
        ForEach-Object {
            if ($_.Version -gt 3.5) {
                $boolNetFXVersionOK = $true
            }
        }

        If (!$boolNetFXVersionOK) {

            Write-Error -Message "This function requires Microsoft .NET Framework version 3.5 or greater."

            Break
        }

        $boolNetFXVersionOK = $null

        $Results=@()

        $StartTime = Get-Date

        $WarningColor = 'Yellow'

        $CriticalColor = 'Red'

        $i=0
}

Process {

    If (!$(Test-Connection -ComputerName $ComputerName -Quiet -Count 1)) {

        Write-Error "Computer $($ComputerName) is unreachable."

        Break

    }


    If ($LogsScope -eq 'Classic') {

        $Logs = Get-WinEvent -ComputerName $ComputerName -ListLog * -ErrorAction SilentlyContinue  |  Where-Object { $_.IsClassicLog }

    }
    Else {

        $Logs = Get-WinEvent -ComputerName $ComputerName -ListLog * -ErrorAction SilentlyContinue

    }

    $LogsCount = ($Logs | Measure-Object).Count

    $Logs | ForEach-Object -Process {

        #Checking if Verbose parameter is not set to true
        If ( !$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent ) {

            $StatusText = "Percent completed $PercentCompleted%, currently the log {0} is checked. " -f $($_.LogName).ToString()

            If ($i -eq 0) {

                $PercentCompleted = 0

            }
            Else {

                $PercentCompleted = [math]::Round(($i / $LogsCount) * 100)

            }

            Write-Progress -Activity "Checking logs for $ComputerName" -Status $StatusText -PercentComplete $PercentCompleted

        }

        $LogTimeSpan = 0

        If ( $Scope -eq 'Classic' ) {

            $OldestEventEntryTime = Get-OldestEventTime -ComputerName $ComputerName -LogName $_.LogName.ToString() -Method LogParser -Verbose:($PSBoundParameters['Verbose'] -eq $true)

        }
        Else{

            If ($_.IsClassicLog -and $LogParserInstalled ) {

                $OldestEventEntryTime = Get-OldestEventTime -ComputerName $ComputerName -LogName $_.LogName.ToString() -Method LogParser -Verbose:($PSBoundParameters['Verbose'] -eq $true)

            }
            Else {

                $OldestEventEntryTime = Get-OldestEventTime -ComputerName $ComputerName -LogName $_.LogName.ToString() -Method PowerShell -Verbose:($PSBoundParameters['Verbose'] -eq $true)

            }
        }

        $LogSpanStatus = 'Empty'

        If ($OldestEventEntryTime) {

            $LogTimeSpan = New-TimeSpan -Start $OldestEventEntryTime -End $StartTime

            If ($LogTimeSpan -lt $(New-TimeSpan -Days $WarningLevelDays) -and $LogTimeSpan -gt $(New-TimeSpan -Days $CriticalLevelDays)) {

                $LogSpanStatus = 'Warning'

            }
            elseif ( $LogTimeSpan -lt $(New-TimeSpan -Days $CriticalLevelDays)) {

                $LogSpanStatus = 'Critical'

            }
            else {

                $LogSpanStatus = 'Normal'

            }

        }

        $PropertiesArray =  @{
            ComputerName        = $ComputerName
            LogName             = $_.LogName
            OldestEventTime     = $OldestEventEntryTime
            LogTimeSpan         = $LogTimeSpan
            LogSpanStatus       = $LogSpanStatus
        }

        $Result = New-Object -TypeName PSObject -Property $PropertiesArray

        Write-Verbose -Message $Result

        If (($Result.OldestEventTime) -or  !$ExcludeEmptyLogs) {

            $Results+=$Result
        }

        $i++

    }

}

End {

        $Results | ForEach-Object -Process {

        If ( $ColorOutput) {

            if ($_.LogSpanStatus -eq 'Critical' -and $ColourOutput) {

               Write-ColorOutput -ForeGroundColor $CriticalColor -OutputData $_

            }
            elseif ($_.LogSpanStatus -eq 'Warning' -and $ColourOutput) {

                Write-ColorOutput -ForeGroundColor $WarningColor -OutputData $_

            }
            else {

                Write-Output -InputObject $_

            }
        }

        }
        Else {

            Return $Results

        }

}

}

Function Get-OldestEventTime {

[CmdletBinding()]

param (

    [parameter(mandatory=$false,Position=0)]
    [String]$ComputerName='localhost',

    [parameter(mandatory=$false,Position=1,ValueFromPipeline=$false)]
    [String]$LogName,

    [parameter(mandatory=$false,Position=2)]
    [ValidateSet('LogParser','PowerShell')]
    [String]$Method='LogParser'

)

begin {

    If ( $Method -eq 'LogParser' ) {

        #Code was partially generated by Log Parser Studio tool

        $LogQuery = New-Object -ComObject "MSUtil.LogQuery"

        $InputFormat = New-Object -ComObject "MSUtil.LogQuery.EventLogInputFormat"


        $InputFormat.fullText=1
        $InputFormat.resolveSIDs=0
        $InputFormat.formatMsg=1
        $InputFormat.msgErrorMode="MSG"
        $InputFormat.fullEventCode=0
        $InputFormat.direction="FW"
        $InputFormat.stringsSep="|"
        $InputFormat.binaryFormat="PRINT"
        $InputFormat.ignoreMessageErrors=0

        If ( $ComputerName -eq 'localhost' ) {

            [String]$LogToQuery = $LogName

        }
        Else {

            [String]$LogToQuery = '\\' + $ComputerName + '\' + $LogName

        }

        $SQLQuery = "SELECT TOP 1 TimeGenerated FROM '{0}' ORDER BY TimeGenerated ASC" -f $LogToQuery

        Write-Verbose -Message "Query used for logparser: $SQLQuery"

    }

}

process {

    try{

        If ($Method -eq 'LogParser') {

            $rtnVal = $LogQuery.Execute($SQLQuery, $InputFormat)

            do{
                $lp_return = @{}

                $log_entry = $rtnVal.getrecord()

                $lp_return.add('TimeGenerated',[datetime]$log_entry.getvalue("TimeGenerated"))

                $rtnVal.movenext()


            } while ($rtnVal.atend() -eq $false)

        }
        Else {

            $lp_return = @{}

            $Oldest = Get-WinEvent -ComputerName $ComputerName -LogName $Logname -ErrorAction SilentlyContinue | Select-Object -Last 1 -ErrorAction SilentlyContinue

            If ($Oldest) {

                $lp_return.add('TimeGenerated', $Oldest.TimeCreated  )

            }
        }


        $Result = New-Object -TypeName PSObject -Property $lp_return

        Write-Verbose -Message "The oldest event log entry in log $Logname is dated on $Result.TimeGenerated"

    }

    catch {

        Write-Verbose -Message "The log $LogName on $ComputerName is unavailable or empty."

    }


}

end {

    Return $Result.TimeGenerated

}

}


function Write-ColorOutput {

param (


    [parameter(mandatory=$true,Position=0)]
    [String]$ForeGroundColor,

    [parameter(mandatory=$true,Position=1)]
    $OutputData

)

    #Source
    #http://stackoverflow.com/questions/4647756/is-there-a-way-to-specify-a-font-color-when-using-write-output

    #Modified by Wojciech Sciesinski

    # save the current color
    $fc = $host.UI.RawUI.ForegroundColor

    # set the new color
    $host.UI.RawUI.ForegroundColor = $ForegroundColor

    # Write output
    Write-Output -InputObject $OutputData

    # restore the original color
    $host.UI.RawUI.ForegroundColor = $fc

}
