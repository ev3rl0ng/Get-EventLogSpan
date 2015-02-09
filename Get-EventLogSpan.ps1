Function Get-EventLogSpan {

<#
	.SYNOPSIS
	Function for veryfing Windows events log span - how old is the oldest available log event entry
	
	.DESCRIPTION
	As a result for this function you receive range between the oldest event log entry and the when the script was started - eg. 3 days 7 hours ...
  
	.PARAMETER ComputerName
	Computer which need to queried. 
  
	.PARAMETER LogsScope
	You can select between Classic and All - default is Classic,
    For basic installation Windows Server 2008 R2 these logs are: Application, Security, Setup, System, Forwarded Events. 
    All logs means that also "Applications and Services Logs" will be included - query all logs can be time consuming because LogParse can't be used for it.
  
	.ExcludeEmptyLogs
    By default empty logs are excluded from results.


    .WarningLevelDays
    The amount od days for which log can be marked in results as Warning if the span of logs is smaller than warning level.
    Default warning level is set to 30 days.

    
    .CriticalLevelDays
    The amount od days for which log can be marked in results as Critical if the span of logs is smaller than warning level.
    Default warning level is set to 7 days.

    
    .OutputDirection
    Default output direction is Console - result is returned as PowerShell object so can be simply redirected to pipe.
    In the future output to HTML file and email will be implemented also.

    
    .ColourOutput
    By default output will be displayed using different colors for Critical and Warning log span levels. 
    In thi moment colors can be changed only by editing source code.

	
	.LogParserInstalled
	Use this (set to $false) if Log Parser is not installed, the native PowerShell cmdlet will be used to parse all event logs. 
	If set to $false log parsing will be much slower - due to PowerShell  for big event logs 

     
	.EXAMPLE
    
    Get-EventLogSpan -ComputerName SRV087.MYDOMAIN.LOCAL -Scope Classic -WarningLevelDays 14


    .LINK
    https://github.com/it-praktyk/Get-EventLogSpan



      
	AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
		
	KEYWORDS: Windows, PowerShell, EventLogs
   
	BASE REPOSITORY: https://github.com/it-praktyk/Get-EventLogSpan

	VERSION HISTORY
	0.3.0 - 2015-01-20 - first version published on GitHub
    0.4.0 - 2015-01-21 - output updated - now include timespan, warning and critical levels added as parameters, output can be coloured
	0.5.0 - 2015-01-25 - checking not classic logs corrected,progress indicator added, lots improvments added 
	0.5.1 - 2015-01-26 - checking oldest log corrected for remote computers
	0.5.2 - 2015-01-26 - checking using Log Parser corrected, output for status for empty logs corrected
    0.5.3 - 2015-02-05 - double quoute to single quote changed for static strings, minor updates
    0.6.0 - 2015-02-06 - check if .Net Framework 3.5 is installed - needed for Get-WinEvent cmdlet
    0.6.1 - 2015-02-09 - help updated

    TODO
    - information that script need be running as administrator
    - HTML output need to be implemented
    - email output need to bi implemented (?)
    - check if .Net 3.5 is installed - http://blog.smoothfriction.nl/archive/2011/01/18/powershell-detecting-installed-net-versions.aspx - needed by Get-WinEvent

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

	#HTML Output is not implemented yet	
    [parameter(mandatory=$false,Position=5)]
    [ValidateSet('Console','HTML')]
    [String]$OutputDirection='Console',

    [parameter(mandatory=$false,Position=6)]
    [Bool]$ColourOutput=$true,
	
	[parameter(mandatory=$false,Position=7)]
	[Bool]$LogParserInstalled=$true

)

Begin {

		#Set-StrictMode -Version 2		

	    $Results=@()

        $StartTime = Get-Date

        $WarningColor = 'Yellow'
        
        $CriticalColor = 'Red' 
		
		$i=0

        

}

Process {

    If (Test-Key "HKLM:\Software\Microsoft\NET Framework Setup\NDP\v3.5" "Install") {
	
	    if ($LogsScope -eq 'Classic') {

		    $Logs = Get-WinEvent -ComputerName $ComputerName -ListLog * -ErrorAction SilentlyContinue  |  Where-Object { $_.IsClassicLog }
		
	    }
	    Else {
	
		    $Logs = Get-WinEvent -ComputerName $ComputerName -ListLog * -ErrorAction SilentlyContinue
	
	    }

    }

    Else {

        Write-Error "This function requires Microsoft .NET Framework version 3.5 or greater."

        Break

    }
	
	    $LogsCount = ($Logs | Measure-Object).Count
	
	$Logs | ForEach-Object -Process {
	
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

        If ($OldestEventEntryTime -ne $null ) {

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

       
		
		$hash =  @{ 
			ComputerName     	= $ComputerName
			LogName				= $_.LogName
			OldestEventTime		= $OldestEventEntryTime
            LogTimeSpan			= $LogTimeSpan
            LogSpanStatus		= $LogSpanStatus
		} 
                
		$Result = New-Object PSObject -Property $hash 
				
		Write-Verbose $Result
		
		If (($Result.OldestEventTime -ne $null) -or  !$ExcludeEmptyLogs) {
				
			$Results+=$Result
		}
		
		$i++
			
		#Checking if Verbose parameter is set to true
		If ( !$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent ) {
		
			$PercentCompleted = ($i / $LogsCount) * 100
		
			Write-Progress -Activity "Checking logs for $ComputerName" -Status "Percent completed $PercentCompleted%" -PercentComplete $(  $PercentCompleted  )
						
		}
		
	}
	
}

End {

    If ($OutputDirection -eq 'Console') {

        $Results | ForEach-Object -Process {

        
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
	
		#Write-Host "Sorry, HTML output is not implemented yet. If needed please use 
		
		#$Results | ConvertTo-HTML | OUT-FILE ".\result.htm"
		
	
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

	#Write-host "Executing query for $LogName on $ComputerName"

	If ( $Method -eq 'LogParser' ) {

		#Code was partially generated by Log Parser Studio tool
	
		$LogQuery = New-Object -ComObject "MSUtil.LogQuery"
	
		$InputFormat = New-Object -ComObject "MSUtil.LogQuery.EventLogInputFormat"

		$OutputFormat = New-Object -ComObject "MSUtil.LogQuery.NativeOutputFormat"
	
		$InputFormat.fullText=1
		$InputFormat.resolveSIDs=0
		$InputFormat.formatMsg=1
		$InputFormat.msgErrorMode="MSG"
		$InputFormat.fullEventCode=0
		$InputFormat.direction="FW"
		$InputFormat.stringsSep="|"
		$InputFormat.binaryFormat="PRINT"
		$InputFormat.ignoreMessageErrors=0

		$SQLQuery = "SELECT TOP 1 TimeGenerated FROM '{1}' ORDER BY TimeGenerated ASC" -f $ComputerName,$LogName
		
		Write-Verbose $SQLQuery
		
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
	
	}

	catch {
		
		Write-Verbose "The log $LogName on $ComputerName is unavailable or empty."

	}
	
	Finally {
			
			$Result = New-Object PSObject -Property $lp_return
			
			Write-Verbose $Result
	
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
    Write-Output $OutputData

    # restore the original color
    $host.UI.RawUI.ForegroundColor = $fc

}


function Test-Key([string]$path, [string]$key) {

    #Source
    #http://blog.smoothfriction.nl/archive/2011/01/18/powershell-detecting-installed-net-versions.aspx

    if(!(Test-Path $path)) { return $false }
    if ((Get-ItemProperty $path).$key -eq $null) { return $false }
    return $true
}