Function Get-EventLogSpan {

#Requires -Version 2.0 
[CmdletBinding()] 


<#
	.SYNOPSIS
	Function for veryfing Windows events log span
  
	.PARAMETER ComputerName

  
	.PARAMETER LogsScope
  
	.ExcludeEmptyLogs

     
	.EXAMPLE

      
	AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
		
	KEYWORDS: Windows, PowerShell, EventLogs
   
	BASEREPOSITORY: https://github.com/it-praktyk/Get-EventLogSpan

	VERSION HISTORY
	0.3 - 2015-01-20 - first version published on GitHub

   #>
   
#>

param (
	[parameter(mandatory=$false,Position=0)]
	[String]$ComputerName="localhost",
	
	[parameter(mandatory=$false,Position=1)]
	[ValidateSet("All","Classic")]
	[String]$LogsScope="Classic",
	
	[parameter(mandatory=$false,Position=2)]
	[Bool]$ExcludeEmptyLogs=$true

)

Begin {

	    $Results=@()

}

Process {
	
	if ($LogsScope -eq "Classic") {

		$Logs = Get-WinEvent -ListLog * |  Where-Object { $_.IsClassicLog }
		
	}
	Else {
	
		$Logs = Get-WinEvent -ListLog *
	
	}
	
	$Logs | ForEach {
	
		$OldestEventEntryTime = Get-OldestEventTime -LogName $_.LogName.ToString()
		
		$hash =  @{ 
			ComputerName     	= $ComputerName
			LogName				= $_.LogName
			OldestEventTime		= $OldestEventEntryTime
		} 
                
		$Result = New-Object PSObject -Property $hash 
				
		Write-Verbose $Result
		
		If (($Result.OldestEventTime -ne $null) -or  !$ExcludeEmptyLogs) {
				
			$Results+=$Result
		}
		
	}
	
}

End {

	Return $Results

}

}

Function Get-OldestEventTime {

[CmdletBinding()] 

param (

	[parameter(mandatory=$false,Position=0)]
	[String]$ComputerName="localhost",

	[parameter(mandatory=$false,Position=1,ValueFromPipeline=$false)]
	[String]$LogName

)

begin {

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
	

	#Write-Verbose "Executing query for $LogName on $ComputerName"

	$SQLQuery = "SELECT TOP 1 TimeGenerated FROM '{1}' ORDER BY TimeGenerated ASC" -f $ComputerName,$LogName
	
	$rtnVal = $LogQuery.Execute($SQLQuery, $InputFormat)
	
}

process {
	
	try{
		do{
			$lp_return = @{}
	
			$log_entry = $rtnVal.getrecord()

			$lp_return.add("TimeGenerated",[datetime]$log_entry.getvalue("TimeGenerated"))

			$rtnVal.movenext()
	

		} while ($rtnVal.atend() -eq $false)
	
	}

	catch {
	
		$lp_return.add("TimeGenerated",$null)
		
		Write-Verbose "The log $LogName on $ComputerName is unavailable or empty."

	}
	
	Finally {
	
			$Result = New-Object PSObject -Property $lp_return
	
	}
	
}

end {

	$Result = New-Object PSObject -Property $lp_return

	Return $Result.TimeGenerated

}

}

