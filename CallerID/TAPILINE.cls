VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAPILINE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************
'* Module:  TAPILINE.cls
'* Purpose: The main TAPI wrapper class
'*
'* Author:  Sumanta Ray
'* Email:   chandsumant@hotmail.com
'* Date:    14/07/2005
'* Copyright (c) 2003-2005 Sumanta Ray. All rights reserved
'***********************************************************************
'* Note: Many of the portions of this class was taken from
'*       Ray Mercer's code at http://www.shrinkwrapvb.com.
'*       Also some portion is adopted from Steven Frare's
'*       code on his web site. Hence i pay due respect and
'*       credit to them.
'***********************************************************************


'Initialization, error & version control variables
Private m_hLineApp As Long 'handle to TAPI
Private m_LowAPI As Long 'default TAPI 1.3 (&H00010003)
Private m_HiAPI As Long 'default TAPI 3.0 (&H00030000)
Private m_LastTAPIError As Long

'line selection variables
Private m_numLines As Long
Private m_CurLineID As Long

'devcaps variables
Private m_linecaps As LINEDEVCAPS
Private m_LineName As String

'Call-state and asynch function state variables
Private m_hLine As Long
Private m_hCall As Long
'line params
Private m_BearerModes As Long
Private m_MediaModes As Long

'CID vars
Private CIDNumber As String
Private CIDName As String

'Events
Event OnCallerID(ByVal CID As String)
Event OnCallerName(ByVal CNAME As String)
Event OnError(ByVal msg As String, ByVal source As String)
Event OnDebug(ByVal msg As String)

Private Sub Class_Initialize()
    Debug.Print "class init"
    LogFile = App.Path & "\" & Format(Date, "mm-dd-yy") & ".log"
    'default TAPI 1.3 - 3.0
    WriteToLog " "
    WriteToLog "*********************START*********************"
    WriteToLog "TAPICallerID Log --> Started"
    m_LowAPI = &H10003
    m_HiAPI = &H30000
End Sub

Public Property Get LastError() As Long
    LastError = m_LastTAPIError
    'reset error value after access
    m_LastTAPIError = 0&
End Property

Public Function InitLines() As Boolean

    Dim rc As Long
    Dim line As Long
    Dim lip As LINEINITIALIZEEXPARAMS
    Dim lxid As LINEEXTENSIONID
        
    'init params
    lip.dwTotalSize = Len(lip)
    lip.dwNeededSize = Len(lip)
    lip.dwUsedSize = Len(lip)
    lip.dwOptions = LINEINITIALIZEEXOPTION_USEHIDDENWINDOW
    
    RaiseEvent OnDebug("Initializing TAPI Lines...")
    'initialize TAPI lines and get handle
    rc = lineInitializeEx(m_hLineApp, App.hInstance, AddressOf LineCallbackProc, App.Title, _
                                m_numLines, m_HiAPI, lip)
    If rc <> TAPI_SUCCESS Then
        m_LastTAPIError = rc
        ReDim LineNames(0)
        LineNames(0) = "No TAPI Line device found"
        RaiseEvent OnError(GetLineErrString(rc), "lineInitializeEx")
        WriteToLog "Error:::" & GetLineErrString(rc) & " @ lineInitializeEx"
        Exit Function
    Else
        If m_numLines < 1 Then Exit Function 'no lines!
        ReDim LineNames(m_numLines - 1)
        RaiseEvent OnDebug("Enumerating TAPI Lines...")
        WriteToLog "Enumerating TAPI Lines..."
        For line = 0 To m_numLines - 1
            If GetLineDevCaps(line) = False Then
                LineNames(line) = "Line - " & line & " (Error retrieving Line Name)"
            End If
            WriteToLog "        Line " & line & " => " & LineNames(line)
            RaiseEvent OnDebug("  Line " & line & " => " & LineNames(line))
        Next
    End If
    
    InitLines = True 'return success
    RaiseEvent OnDebug("Line Initalized Successfully")
    WriteToLog "Line Initalized Successfully"
End Function

Public Property Get NumLines() As Long
    NumLines = m_numLines
End Property

Public Property Get CurrentLineID() As Long
    CurrentLineID = m_CurLineID
End Property

Public Property Let CurrentLineID(ByVal id As Long)
    If id < 0 Then id = 0
    If id > m_numLines - 1 Then id = (m_numLines - 1)
    
    m_CurLineID = id
    m_LineName = LineNames(m_CurLineID)
    RaiseEvent OnDebug("Current Line ID = " & m_CurLineID & " (" & LineNames(m_CurLineID) & ")")
    WriteToLog "Current Line ID = " & m_CurLineID & " (" & LineNames(m_CurLineID) & ")"
End Property

Public Property Get LineName() As String
    LineName = m_LineName
End Property

Public Property Get LastCallerID() As String
    LastCallerID = CIDNumber
End Property

Public Property Get LastCallerName() As String
    LastCallerName = CIDName
End Property

Private Function GetLineDevCaps(ByVal LineID As Integer) As Boolean
    Dim rc As Long
    WriteToLog "Verifying Line device capability..."
    'init required fields in devcaps UDT
    m_linecaps.dwTotalSize = Len(m_linecaps)
    'request TAPI to fill UDT with device capability info
    rc = lineGetDevCaps(m_hLineApp, LineID, TAPI_CURRENT_VERSION, _
                        0&, m_linecaps)
    If rc <> TAPI_SUCCESS Then
        m_LastTAPIError = rc
        RaiseEvent OnError(GetLineErrString(rc), "lineGetDevCaps")
        WriteToLog "Error:::" & GetLineErrString(rc) & " @ lineGetDevCaps"
        Exit Function
    End If
    
    'store UDT info in class local variables
    m_LineName = StructPtrToString(VarPtr(m_linecaps), m_linecaps.dwLineNameOffset, m_linecaps.dwLineNameSize)
    m_BearerModes = m_linecaps.dwBearerModes
    m_MediaModes = m_linecaps.dwMediaModes
    LineNames(LineID) = m_LineName
     'return success
    GetLineDevCaps = True
    WriteToLog "lineGetDevCaps - > Success"
End Function

Public Property Get LineSupportsVoiceCalls() As Boolean
    If m_BearerModes And LINEBEARERMODE_VOICE Then
        If m_MediaModes And LINEMEDIAMODE_INTERACTIVEVOICE Then
            LineSupportsVoiceCalls = True
        End If
    End If
End Property

Public Function OpenLine(Optional ByVal privileges As Long, _
                        Optional ByVal mediamodes As Long) As Boolean
    'returns false on error or if a line is already open
    'if a TAPI error occurs the error number will be stored in m_LastTAPIerror
    If m_hLine <> 0 Then
        RaiseEvent OnError("OpenLine invoked recursively", "OpenLine")
        WriteToLog "Error::: OpenLine invoked recursively" & " @ OpenLine"
        Exit Function
    End If
    Dim rc As Long
    RaiseEvent OnDebug("Opening Line " & m_CurLineID)
    WriteToLog "Opening Line " & m_CurLineID
    rc = lineOpen(m_hLineApp, _
                    m_CurLineID, _
                    m_hLine, _
                    TAPI_CURRENT_VERSION, _
                    0&, _
                    ByVal ObjPtr(Me), _
                    privileges, _
                    mediamodes, _
                    ByVal 0&)
    If rc <> TAPI_SUCCESS Then
        m_LastTAPIError = rc
        RaiseEvent OnError(GetLineErrString(rc), "lineOpen")
        RaiseEvent OnDebug("Failed to open Line " & m_CurLineID)
        WriteToLog "Error:::" & GetLineErrString(rc) & " @ lineOpen"
        Exit Function
    End If
    
       'setting line status messages
    rc = lineSetStatusMessages(m_hLine, m_linecaps.dwLineStates, 0)
    If rc <> 0 Then
        m_LastTAPIError = rc
        RaiseEvent OnError(GetLineErrString(rc), "lineSetStatusMessages")
        WriteToLog "Error:::" & GetLineErrString(rc) & " @ lineSetStatusMessages"
        Exit Function
    End If
    'Debug.Assert m_hLine
    OpenLine = True
    RaiseEvent OnDebug("Line " & m_CurLineID & " opened sucessfully")
    WriteToLog "Line " & m_CurLineID & " opened sucessfully"
End Function

Public Function CloseLine() As Boolean
    CloseLine = False
    Dim rc As Long
    RaiseEvent OnDebug("Closing Line " & m_CurLineID)
    WriteToLog "Closing Line " & m_CurLineID
    If m_hLine = 0 Then
        RaiseEvent OnError("No line to close", "CloseLine")
        WriteToLog "No Line to Close" & " @ CloseLine"
        Exit Function
    End If
    
    rc = lineClose(m_hLine)
    Debug.Print "Line Closed"
    m_hLine = 0
    If rc <> TAPI_SUCCESS Then
        m_LastTAPIError = rc
        RaiseEvent OnError(GetLineErrString(rc), "lineClose")
        RaiseEvent OnDebug("Failed to close Line " & m_CurLineID)
        WriteToLog "Error:::" & GetLineErrString(rc) & " @ lineClose"
        Exit Function
    End If
    
    CloseLine = True
    RaiseEvent OnDebug("Line " & m_CurLineID & " Closed Successfully")
    WriteToLog "Line " & m_CurLineID & " Closed Successfully"
End Function

Private Function StructPtrToString(ByVal ptrStruct As Long, ByVal offset As Long, ByVal length As Long) As String
    Dim MyBuffer() As Byte

    If length < 1 Then Exit Function

    If offset Then '
        ReDim MyBuffer(0 To length - 1)
        CopyMemory MyBuffer(0), ByVal ptrStruct + offset, length
        StructPtrToString = StrConv(MyBuffer, vbUnicode)
    End If

End Function

Private Sub Class_Terminate()

    Debug.Print "class term"
    Dim rc As Long
    If m_hCall <> 0 Then
        rc = lineDeallocateCall(m_hCall)
        If rc <> TAPI_SUCCESS Then
            m_LastTAPIError = rc
            RaiseEvent OnError(GetLineErrString(rc), "lineDeallocateCall")
            WriteToLog "Error:::" & GetLineErrString(rc) & " @ lineDeallocateCall"
        Else
            m_hCall = 0
        End If
    End If
    
    If m_hLine <> 0 Then
        If Not CloseLine Then
            RaiseEvent OnError("Failed to close the line", "Class_Terminate")
        End If
    End If
    
    If m_hLineApp <> 0 Then
        rc = lineShutdown(m_hLineApp)
        If rc <> TAPI_SUCCESS Then
            m_LastTAPIError = rc
            RaiseEvent OnError(GetLineErrString(rc), "lineShutdown")
            WriteToLog "Error:::" & GetLineErrString(rc) & " @ lineShutdown"
        Else
            m_hLineApp = 0
        End If
    End If
    WriteToLog "*********************END*********************"
    WriteToLog " "
End Sub

Friend Sub LineProcHandler(ByVal hDevice As Long, _
                                        ByVal dwInstance As Long, _
                                        ByVal dwMsg As Long, _
                                        ByVal dwParam1 As Long, _
                                        ByVal dwParam2 As Long, _
                                        ByVal dwParam3 As Long)
'Handle callbacks here in the class itself via ITapiCallbackSink
    Dim rc As Long
    Debug.Print "Entering LineProcHandler"
    WriteToLog "Entering LineProcHandler"
    Select Case dwMsg
        
        Case LINE_CALLINFO
            Debug.Print "LINE_CALLINFO"
            If dwParam1 = LINECALLINFOSTATE_CALLERID Then
                'Get the Caller ID info here
                WriteToLog "Call Information Received..."
                rc = GetCallerInfo(hDevice)
                Debug.Print "rc=" & rc
                If rc <> 0 Then
                    m_LastTAPIError = rc
                    RaiseEvent OnError("Failed to retrieve call information", "GetCallerInfo")
                    WriteToLog "Error:::" & GetLineErrString(rc) & " @ GetCallerInfo"
                End If
            Else
                Debug.Print "LINE_CALLINFO -> " & CStr(dwParam1)
            End If
        
        Case LINE_CALLSTATE
            Select Case dwParam1
                              
                Case LINECALLSTATE_IDLE
                    'no call exists - the line is idle
                    Debug.Print "LINECALLSTATE_IDLE"
                    RaiseEvent OnDebug("Call State is IDLE")
                    WriteToLog "Call State is IDLE"
                    If m_hCall <> 0 Then
                        Debug.Print "Deallocating Call"
                        RaiseEvent OnDebug("Deallocating Call...")
                        WriteToLog "Deallocating Call..."
                        rc = lineDeallocateCall(m_hCall)
                        If rc <> 0 Then
                            m_LastTAPIError = rc
                            RaiseEvent OnError(GetLineErrString(rc), "lineDeallocateCall")
                            WriteToLog "Error:::" & GetLineErrString(rc) & " @ lineDeallocateCall"
                        End If
                        m_hCall = 0
                        RaiseEvent OnDebug("Line Deallocated")
                        WriteToLog "Line Deallocated"
                    End If
                
                Case LINECALLSTATE_OFFERING
                    'init caller id vars at offering state
                    CIDName = "Unknown"
                    CIDNumber = "Unknown"
                    m_hCall = hDevice
                    RaiseEvent OnDebug("Offering Call..")
                    WriteToLog "Offering Call.."
                Case Else
                    Debug.Print "CallSTATE: " & Hex(dwParam1)
            End Select
        
    End Select

End Sub

Private Function GetCallerInfo(ByVal hCall As Long) As Long
    Dim lpCallInfo As LINECALLINFO
    Dim nErr As Long
    Dim sCallerID As String
    Dim sCallerName As String
    Dim lStart As Long
    Dim lLength As Long
    Dim lLoop As Long

    On Error GoTo eh

    lpCallInfo.dwTotalSize = Len(lpCallInfo)
    nErr = lineGetCallInfo(hCall, lpCallInfo)
        
    'We'll bail here if the 2k HACK isn't big enough
    If nErr <> 0 Then
        GetCallerInfo = nErr
        Exit Function
    End If
    'Check the LINECALLPARTYID_ Constant to see if we have good info.
    'If (lpCallInfo.dwCallerIDFlags And LINECALLPARTYID_ADDRESS) <> False Then
    If (lpCallInfo.dwCallerIDFlags And (LINECALLPARTYID_ADDRESS Or LINECALLPARTYID_PARTIAL)) <> False Then
    
        'Find the true length of the lpCallInfo UDT (i.e. subtract the bBytes added)
        lStart = Len(lpCallInfo) - UBound(lpCallInfo.bBytes())
        'With the actual length we can now index into the appended info.
        lStart = lpCallInfo.dwCallerIDOffset - lStart + 1
        lLength = lpCallInfo.dwCallerIDSize
        'look if anything is there in the added hack
        For lLoop = 0 To lLength
            If lpCallInfo.bBytes(lStart + lLoop) = 0 Then Exit For
            sCallerID = sCallerID & CStr(Chr(lpCallInfo.bBytes(lStart + lLoop)))
        Next
    
        CIDNumber = sCallerID
    
    Else
        CIDNumber = "Unknown"
    End If
    
    RaiseEvent OnCallerID(CIDNumber)
    
    'This is an exact repeat of above, except for the name not the number.
    If (lpCallInfo.dwCallerIDFlags And LINECALLPARTYID_NAME) <> False Then
        lStart = Len(lpCallInfo) - UBound(lpCallInfo.bBytes())
        lStart = lpCallInfo.dwCallerIDNameOffset - lStart + 1
        lLength = lpCallInfo.dwCallerIDNameSize
    
        For lLoop = 0 To lLength
            If lpCallInfo.bBytes(lStart + lLoop) = 0 Then Exit For
            sCallerName = sCallerName & CStr(Chr(lpCallInfo.bBytes(lStart + lLoop)))
        Next
    
        CIDName = sCallerName
    Else
        CIDName = "Unknown"
    End If
    
    RaiseEvent OnCallerName(CIDName)
    
    Debug.Print "Caller ID = " & CIDNumber
    Debug.Print "Caller Name =" & CIDName
    RaiseEvent OnDebug("Call Information Received...")
    GetCallerInfo = 0
    Exit Function
        
eh:
    Debug.Print Err.Description & " @ GetCallInfo"
    GetCallerInfo = Err.Number
End Function

