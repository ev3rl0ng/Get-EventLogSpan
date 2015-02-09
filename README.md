# PowerShell function Get-EventLogSpan

is intended to check how old is the oldest event entry in event log.

Event logs can be parsed using Log Parser 2.2 or Windows PowerShell Get-WinEvent cmdlet from local or remote machine.


As the function result is returned PowerShell object which contain data like below
LogName				: Application
LogSpanStatus		: Normal
ComputerName		: Localhost
OldestEventTime		: 2013-12-06 00:45:13
LogTimeSpan			: 430.23:02:55.7739014

Function contributed also to Technet Gallery https://gallery.technet.microsoft.com/PowerShell-Function-to-2ea8205a