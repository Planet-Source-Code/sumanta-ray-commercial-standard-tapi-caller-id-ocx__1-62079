Attribute VB_Name = "General"
'*************************************************************
'* Module:  General.bas
'* Purpose: Common routines and APIs
'*
'* Author:  Sumanta Ray
'* Email:   chandsumant@hotmail.com
'* Date:    14/07/2005
'* Copyright (c) 2003-2005 Sumanta Ray. All rights reserved
'*************************************************************

Public LineNames() As String 'for storing the TAPI Line Names
Public LogFile As String

Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
                                    (dest As Any, src As Any, ByVal length As Long)

Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Wparam As Long, ByVal lParam As Long) As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function WriteToLog(ByVal str As String)
    On Error GoTo eh
    Dim fh
    fh = FreeFile
    If LogFile = "" Then
        LogFile = App.Path & "\" & Format(Date, "mm-dd-yy") & ".log"
    End If
    Open LogFile For Append Shared As #fh
        Print #fh, Now & " " & str
    Close #fh
    Exit Function
eh:
    Close #fh
End Function

