<div align="center">

## Advanced Windows NT Eventlogger


</div>

### Description

A VB6 class file that allows to easily write messages - even with more than one parameter (best used with additional messagetable in resource DLL) to the windows eventlog.
 
### More Info
 
type and content of a message to be logged to the eventlog


<span>             |<span>
---                |---
**Submitted On**   |2003-04-10 11:14:04
**By**             |[Holger Sauer](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/holger-sauer.md)
**Level**          |Intermediate
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Advanced\_W1571864102003\.zip](https://github.com/Planet-Source-Code/holger-sauer-advanced-windows-nt-eventlogger__1-44655/archive/master.zip)

### API Declarations

```
Private Declare Function RegisterEventSource Lib "advapi32.dll" Alias "RegisterEventSourceA" (ByVal lpUNCServerName As String, ByVal lpSourceName As String) As Long
Private Declare Function DeregisterEventSource Lib "advapi32.dll" (ByVal hEventLog As Long) As Long
Private Declare Function ReportEvent Lib "advapi32.dll" Alias "ReportEventA" (ByVal hEventLog As Long, ByVal wType As Integer, ByVal wCategory As Integer, ByVal dwEventID As Long, ByVal lpUserSid As Any, ByVal wNumStrings As Integer, ByVal dwDataSize As Long, plpStrings As Long, lpRawData As Any) As Boolean
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Const EVENTLOG_SUCCESS = 0
Private Const EVENTLOG_ERROR_TYPE = &H1
Private Const EVENTLOG_WARNING_TYPE = &H2
Private Const EVENTLOG_INFORMATION_TYPE = &H4
Private Const EVENTLOG_AUDIT_SUCCESS = 8
Private Const EVENTLOG_AUDIT_FAILURE = 10
```





