VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventLogger"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
'*************************************************************************************************
'EventLogger
'Visual Basic class for logging events into the Windows NT eventlog
'written on 2003-04-02 by Holger Sauer (http://www.h-sauer.de)
'This file may be used and distributes under the conditions of the GNU GPL
'*************************************************************************************************

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

Private Const Const_Program_ProgramName = "EventLogTest"
    
Enum ErrorType
    SuccessMsg = 1
    InfoMsg = 2
    WarningMsg = 3
    ErrorMsg = 4
End Enum
    
Private Function UBoundEx(varArray As Variant) As Variant
On Error GoTo lblError:
  UBoundEx = UBound(varArray)
Exit Function
lblError:
Err.Clear
  UBoundEx = -1
End Function
    
Public Sub WriteLog(lngType As ErrorType, strMsg As String)
  Select Case lngType
    Case SuccessMsg
        Call LogNTEvent(strMsg, EVENTLOG_SUCCESS, 0)
    Case InfoMsg
        Call LogNTEvent(strMsg, EVENTLOG_INFORMATION_TYPE, &H40000001)
    Case WarningMsg
        Call LogNTEvent(strMsg, EVENTLOG_WARNING_TYPE, &H80000002)
    Case ErrorMsg
        Call LogNTEvent(strMsg, EVENTLOG_ERROR_TYPE, &HC0000003)
  End Select
End Sub

Public Sub WriteLogEx(lngType As Long, lngEventID As Long, strMessages() As String)
  Select Case lngType
    Case ErrorType.SuccessMsg
        Call LogNTEventEx(strMessages, EVENTLOG_SUCCESS, lngEventID)
    Case ErrorType.InfoMsg
        Call LogNTEventEx(strMessages, EVENTLOG_INFORMATION_TYPE, lngEventID)
    Case ErrorType.WarningMsg
        Call LogNTEventEx(strMessages, EVENTLOG_WARNING_TYPE, lngEventID)
    Case ErrorType.ErrorMsg
        Call LogNTEventEx(strMessages, EVENTLOG_ERROR_TYPE, lngEventID)
  End Select
End Sub

Private Sub LogNTEventEx(strStrings() As String, intLogType As Integer, lngEventID As Long)
    On Error GoTo ErrHandler
    
    Dim intNumStrings As Integer
    Dim intTemp As Integer
    Dim hEventLog As Long
    Dim hMsgs As Long
    Dim strTemp As String
    Dim varTemp As Variant
    Dim lngMessageSize As Long
    Dim arrTemp() As Long
    
    'register program at eventlog
    hEventLog = RegisterEventSource("", Const_Program_ProgramName)
    'are parameter strings existing?
    If (UBoundEx(strStrings) = -1) Then
        intNumStrings = 0
        ReDim arrTemp(0)
        arrTemp(0) = vbNullChar
    Else
        'create array of parameters
        intNumStrings = UBoundEx(strStrings) + 1
        ReDim arrTemp(intNumStrings - 1)
        intTemp = 0
        For Each varTemp In strStrings
            strTemp = varTemp & vbNullChar
            lngMessageSize = Len(strTemp)
            hMsgs = GlobalAlloc(&H40, lngMessageSize)
            CopyMemory ByVal hMsgs, ByVal strTemp, lngMessageSize
            arrTemp(intTemp) = hMsgs
            intTemp = intTemp + 1
        Next
    End If
    'write event to eventlog
    If ReportEvent(hEventLog, intLogType, 0, lngEventID, 0&, intNumStrings, 0, arrTemp(0), 0&) = 0 Then
        'error occured
        'MsgBox GetLastError()
    End If
    GlobalFree (hMsgs)
    'unregister program at eventlog
    DeregisterEventSource (hEventLog)
    
    Exit Sub
ErrHandler:
    'MsgBox Err.Number & ": " & Err.Description
    Err.Clear
End Sub

Private Sub LogNTEvent(sString As String, iLogType As Integer, iEventID As Long)
    On Error GoTo ErrHandler
    
    Dim bRC As Boolean
    Dim iNumStrings As Integer
    Dim hEventLog As Long
    Dim hMsgs As Long
    Dim cbStringSize As Long
    
    hEventLog = RegisterEventSource("", Const_Program_ProgramName)
    cbStringSize = Len(sString) + 1
    hMsgs = GlobalAlloc(&H40, cbStringSize)
    CopyMemory ByVal hMsgs, ByVal sString, cbStringSize
    iNumStrings = 1
    If ReportEvent(hEventLog, iLogType, 0, iEventID, 0&, iNumStrings, cbStringSize, hMsgs, hMsgs) = 0 Then
        'MsgBox GetLastError()
    End If
    GlobalFree (hMsgs)
    DeregisterEventSource (hEventLog)
    
    Exit Sub
ErrHandler:
    'MsgBox Err.Number & ": " & Err.Description
    Err.Clear
End Sub

