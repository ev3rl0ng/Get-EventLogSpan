Function Get-EventLogSpan {

#Requires -Version 2.0 
#Set-strictmode -version 2.0
[CmdletBinding()] 


<#
	.SYNOPSIS
	Function for veryfing Windows events log span
	
	.DESCRIPTION
	As a result for this function you receive range between the oldest event log entry and the when the script was startud
  
	.PARAMETER ComputerName
	Computer which need to queried
  
	.PARAMETER LogsScope
	You c
  
	.ExcludeEmptyLogs


    .WarningLevel

    
    .CriticalLevel

    
    .OutputDirection

    
    .ColourOutput
	
	.LogParserInstalled
	Use this if Log Parser is not installed, the native PowerShell cmdlet will be used to parse all event logs. 
	If set to true log parsing will be very slower for big event logs 

     
	.EXAMPLE

      
	AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
		
	KEYWORDS: Windows, PowerShell, EventLogs
   
	BASE REPOSITORY: https://github.com/it-praktyk/Get-EventLogSpan

	VERSION HISTORY
	0.3 - 2015-01-20 - first version published on GitHub
    0.4 - 2015-01-21 - output updated - now include timespan, warning and critical levels added as parameters, output can be coloured
	0.5 - 2015-01-25 - checking not classic logs corrected,progress indicator added, lots improvments added 

   #>
   
#>

param (
	[parameter(mandatory=$false,Position=0)]
	[String]$ComputerName="localhost",
	
	[parameter(mandatory=$false,Position=1)]
	[ValidateSet("All","Classic")]
	[String]$LogsScope="Classic",
	
	[parameter(mandatory=$false,Position=2)]
	[Bool]$ExcludeEmptyLogs=$true,

    [parameter(mandatory=$false,Position=3)]
    [Int32]$WarningLevelDays=30,

    [parameter(mandatory=$false,Position=4)]
    [Int32]$CriticalLevelDays=7,

	#HTML Output is not implemented yet	
    [parameter(mandatory=$false,Position=5)]
    [ValidateSet("Console","HTML")]
    [String]$OutputDirection="Console",

    [parameter(mandatory=$false,Position=6)]
    [Bool]$ColourOutput=$true,
	
	[parameter (mandatory=$false,Position=7)]
	[Bool]$LogParserInstalled=$true

)

Begin {

	    $Results=@()

        $StartTime = Get-Date

        $WarningColor = "Yellow"
        
        $CriticalColor = "Red" 
		
		$i=0

}

Process {
	
	if ($LogsScope -eq "Classic") {

		$Logs = Get-WinEvent -ComputerName $ComputerName -ListLog *  |  Where-Object { $_.IsClassicLog }
		
	}
	Else {
	
		$Logs = Get-WinEvent -ComputerName $ComputerName -ListLog *
	
	}
	
	$LogsCount = ($Logs | Measure-Object).Count
	
	$Logs | ForEach {
	
		$LogTimeSpan = 0
	
		If ( $Scope -eq "Classic" ) {
	
			$OldestEventEntryTime = Get-OldestEventTime -LogName $_.LogName.ToString() -Method LogParser
			
		}
		Else{
		
			If ($_.IsClassicLog -and $LogParserInstalled ) {
			
				$OldestEventEntryTime = Get-OldestEventTime -LogName $_.LogName.ToString() -Method LogParser
			
			}
			Else {
			
				$OldestEventEntryTime = Get-OldestEventTime -LogName $_.LogName.ToString() -Method PowerShell
			
			}
		}

        If ($OldestEventEntryTime -ne $null ) {

            $LogTimeSpan = New-TimeSpan -Start $OldestEventEntryTime -End $StartTime

            If ($LogTimeSpan -lt $(New-TimeSpan -Days $WarningLevelDays) -and $LogTimeSpan -gt $(New-TimeSpan -Days $CriticalLevelDays)) {

                $LogSpanStatus = "Warning"

            }
            elseif ( $LogTimeSpan -lt $(New-TimeSpan -Days $CriticalLevelDays)) {

                $LogSpanStatus = "Critical"

            }
			elseif ( $LogTimeSpan -eq 0 ) {
			
				$LogSpanStatus = "Empty"
			
			}
            else {

                $LogSpanStatus = "Normal"

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
		If ( !$PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent ) {
		
			Write-Progress -Activity "Checking logs for $ComputerName" -PercentComplete $( ($i / $LogsCount) * 100  )
						
		}
		
	}
	
}

End {

    If ($OutputDirection -eq "Console") {

        $Results | ForEach-Object -Process {

        
            if ($_.LogSpanStatus -eq "Critical" -and $ColourOutput) {
            
               Write-ColorOutput -ForeGroundColor $CriticalColor -OutputData $_
               
            }
            elseif ($_.LogSpanStatus -eq "Warning" -and $ColourOutput) {

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
	[String]$ComputerName="localhost",

	[parameter(mandatory=$false,Position=1,ValueFromPipeline=$false)]
	[String]$LogName,
	
	[parameter(mandatory=$false,Position=2)]
	[ValidateSet("LogParser","PowerShell")]
	[String]$Method="LogParser"

)

begin {

	#Write-host "Executing query for $LogName on $ComputerName"

	If ( $Method -eq "LogParser" ) {

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
		$InputFormat.ignoreMessageErrors=1
	



		$SQLQuery = "SELECT TOP 1 TimeGenerated FROM '{1}' ORDER BY TimeGenerated ASC" -f $ComputerName,$LogName
		
	}
		
}

process {
	
	try{
	
		If ($Method -eq "LogParser") {
		
			I$rtnVal = $LogQuery.Execute($SQLQuery, $InputFormat)
		
			do{
				$lp_return = @{}
	
				$log_entry = $rtnVal.getrecord()

				$lp_return.add("TimeGenerated",[datetime]$log_entry.getvalue("TimeGenerated"))

				$rtnVal.movenext()
	

			} while ($rtnVal.atend() -eq $false)
			
		}
		Else {
		
			$lp_return = @{}
			
			$Oldest = Get-WinEvent -ComputerName $ComputerName -LogName $Logname -ErrorAction SilentlyContinue | Select-Object -Last 1 -ErrorAction SilentlyContinue
			
			If ($Oldest) {
			
				$lp_return.add("TimeGenerated", $Oldest.TimeCreated  )
			
			}
		}
	
	}

	catch {
		
		Write-Verbose "The log $LogName on $ComputerName is unavailable or empty."

	}
	
	Finally {
			
			$Result = New-Object PSObject -Property $lp_return
	
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
