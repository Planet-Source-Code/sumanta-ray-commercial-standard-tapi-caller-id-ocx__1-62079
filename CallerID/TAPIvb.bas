Attribute VB_Name = "TAPIvb"
Option Explicit
'******************************************************************
'* Module:  TAPIvb.bas
'* Purpose: VB Callback Proc for TAPI. Portion taken from
'*          Ray Mercer's code at http://www.shrinkwrapvb.com
'* Author:  Sumanta Ray
'* Email:   chandsumant@hotmail.com
'* Date:    03/07/2004
'* Copyright (c) 1999-2001 Ray Mercer. All rights reserved
'******************************************************************


Public Sub LineCallbackProc(ByVal hDevice As Long, _
                                ByVal dwMsg As Long, _
                                ByVal dwCallbackInstance As Long, _
                                ByVal dwParam1 As Long, _
                                ByVal dwParam2 As Long, _
                                ByVal dwParam3 As Long)
    'the callbackInstance parameter contains a pointer to the TAPILine class
    'this sub just routes all callbacks back to the class for handling there
    Dim PassedObj As TAPILINE
    Dim objTemp As TAPILINE
    Debug.Print "LineCALLBACK : dwCallbackInst = " & dwCallbackInstance
    If dwCallbackInstance <> 0 Then
        'turn pointer into illegal, uncounted reference
        CopyMemory objTemp, dwCallbackInstance, 4
        'Assign to legal reference
        Set PassedObj = objTemp
        'Destroy the illegal reference
        CopyMemory objTemp, 0&, 4
        'use the interface to call back to the class
        PassedObj.LineProcHandler hDevice, dwCallbackInstance, dwMsg, dwParam1, dwParam2, dwParam3
    End If

End Sub



