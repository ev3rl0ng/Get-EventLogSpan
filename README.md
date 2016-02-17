# Get-EventLogSpan
## SYNOPSIS
Function for veryfing Windows events log span - how old is the oldest available log event entry



## SYNTAX
```powershell
Get-EventLogSpan [[-ComputerName] <string>] [[-LogsScope] <string>] [[-ExcludeEmptyLogs] <bool>] [[-WarningLevelDays] <int>] [[-CriticalLevelDays] <int>] [[-ColourOutput] <bool>] [[-LogParserInstalled] <bool>] [<CommonParameters>]                                                                                               
```

## DESCRIPTION
Function for veryfing Windows events log span - how old is the oldest available log event entry.  
As a result for this function you receive range between the oldest event log entry and the newest one.


## PARAMETERS
### -ComputerName &lt;string&gt;
Computer on which logs need to queried.  
By default localhost will be queried.

```
Position?                    0
Accept pipeline input?       false
Parameter set name           (All)
Aliases                      None
Dynamic?                     false
```

### -LogsScope &lt;string&gt;
You can select between Classic and All - default is Classic,
For basic installation Windows Server 2008 R2 these logs are: Application, Security, Setup, System, Forwarded Events.  
All logs means that also "Applications and Services Logs" will be included - query all logs can be time consuming because LogParse can't be used for it.

```
Position?                    1
Accept pipeline input?       false
Parameter set name           (All)
Aliases                      None
Dynamic?                     false
```

### -ExcludeEmptyLogs &lt;bool&gt;
By default empty logs are excluded from results.

```
Position?                    2
Accept pipeline input?       false
Parameter set name           (All)
Aliases                      None
Dynamic?                     false
```

### -WarningLevelDays &lt;int&gt;
The amount od days for which log can be marked in results as Warning if the span of logs is smaller than warning level.  
Default warning level is set to 30 days.

```
Position?                    3
Accept pipeline input?       false
Parameter set name           (All)
Aliases                      None
Dynamic?                     false
```

### -CriticalLevelDays &lt;int&gt;
The amount od days for which log can be marked in results as Critical if the span of logs is smaller than critical level.  
Default critical level is set to 7 days.

```
Position?                    4
Accept pipeline input?       false
Parameter set name           (All)
Aliases                      None
Dynamic?                     false
```


### -LogParserInstalled &lt;bool&gt;
Use this parameter (be default set to $false) if Log Parser is not installed, the native PowerShell cmdlet will be used to parse all event logs.
If set to $false log parsing will be much slower - due to use PowerShell native command to find the oldest log event entry.

```
Position?                    5
Accept pipeline input?       false
Parameter set name           (All)
Aliases                      None
Dynamic?                     false
```


## NOTES

AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net

CONTRIBUTORS:
- Thomas Rhoads, ev3rl0ng[at]gmail[dot]com

KEYWORDS: Windows, PowerShell, EventLogs

VERSION HISTORY  
The versions history you can find [here](VERSIONS.md).

LICENSE  
Copyright (c) 2016 Wojciech Sciesinski  
This function is licensed under The MIT License (MIT)  
Full license text: https://opensource.org/licenses/MIT


## EXAMPLES

### EXAMPLE 1

Get events log span from the localhost

```powershell

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
```
