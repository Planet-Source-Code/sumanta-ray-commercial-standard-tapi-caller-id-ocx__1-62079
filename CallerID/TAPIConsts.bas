Attribute VB_Name = "TAPIConsts"
'*************************************************************
'* Module:  TAPIConsts.bas
'* Purpose: VB32 translation of tapi.h constants
'*          Mostly found from MSDN and Internet
'* Author:  Sumanta Ray
'* Email:   chandsumant@hotmail.com
'* Date:    14/07/2005
'* Copyright (c) 1999-2001 Ray Mercer. All rights reserved
'*************************************************************

Option Explicit

Global Const TAPI_SUCCESS As Long = 0& 'declared for convenience
Global Const TAPI_CURRENT_VERSION = &H20000
'// These constants are mutually exclusive - there's no way to specify more
'// than one at a time (and it doesn't make sense, either) so they're
'// ordinal rather than bits.
Global Const LINEINITIALIZEEXOPTION_USEHIDDENWINDOW   As Long = &H1&        '// TAPI v2.0
Global Const LINEINITIALIZEEXOPTION_USEEVENT    As Long = &H2&         '// TAPI v2.0
Global Const LINEINITIALIZEEXOPTION_USECOMPLETIONPORT    As Long = &H3&         '// TAPI v2.0

'// Messages for Phones and Lines

Global Const LINE_ADDRESSSTATE               As Long = 0&
Global Const LINE_CALLINFO                   As Long = 1&
Global Const LINE_CALLSTATE                  As Long = 2&
Global Const LINE_CLOSE                      As Long = 3&
Global Const LINE_DEVSPECIFIC                As Long = 4&
Global Const LINE_DEVSPECIFICFEATURE         As Long = 5&
Global Const LINE_GATHERDIGITS               As Long = 6&
Global Const LINE_GENERATE                   As Long = 7&
Global Const LINE_LINEDEVSTATE               As Long = 8&
Global Const LINE_MONITORDIGITS              As Long = 9&
Global Const LINE_MONITORMEDIA               As Long = 10&
Global Const LINE_MONITORTONE                As Long = 11&
Global Const LINE_REPLY                      As Long = 12&
Global Const LINE_REQUEST                    As Long = 13&
'Global Const PHONE_BUTTON                    As Long = 14&
'Global Const PHONE_CLOSE                     As Long = 15&
'Global Const PHONE_DEVSPECIFIC               As Long = 16&
'Global Const PHONE_REPLY                     As Long = 17&
'Global Const PHONE_STATE                     As Long = 18&
Global Const LINE_CREATE                     As Long = 19&                      '// TAPI v1.4
'Global Const PHONE_CREATE                    As Long = 20&                      '// TAPI v1.4

'for caller id
Global Const LINECALLINFO_FIXEDSIZE = 296
Global Const LINECALLINFO_VARSIZE = 500
Global Const LINECALLINFO_TOTALSIZE = LINECALLINFO_FIXEDSIZE + LINECALLINFO_VARSIZE
Global Const LINECALLINFOSTATE_CALLERID = 32768

'#if (TAPI_CURRENT_VERSION >= 0x00020000)
Global Const LINE_AGENTSPECIFIC              As Long = 21&                      '// TAPI v2.0
Global Const LINE_AGENTSTATUS                As Long = 22&                      '// TAPI v2.0
Global Const LINE_APPNEWCALL                 As Long = 23&                      '// TAPI v2.0
Global Const LINE_PROXYREQUEST               As Long = 24&                      '// TAPI v2.0
Global Const LINE_REMOVE                     As Long = 25&                      '// TAPI v2.0
Global Const PHONE_REMOVE                    As Long = 26&                      '// TAPI v2.0
'#End If

Public Function GetLineErrString(lParam As Long) As String
'Returns a String description of a TAPI Line Error code
    Dim msg As String
    
    Select Case lParam
        Case LINEERR_ALLOCATED '( = &H80000001)
            msg = "Allocated"
        Case LINEERR_BADDEVICEID '(= &H80000002)
            msg = "Bad Device ID"
        Case LINEERR_BEARERMODEUNAVAIL '(= &H80000003)
            msg = "Bearer Mode Unavail"
        Case LINEERR_CALLUNAVAIL '(= &H80000005)
            msg = "Call UnAvail"
        Case LINEERR_COMPLETIONOVERRUN '(= &H80000006
            msg = "Completion Overrun"
        Case LINEERR_CONFERENCEFULL '(= &H80000007
            msg = "Conference Full"
        Case LINEERR_DIALBILLING '(= &H80000008
            msg = "Dial Billing"
        Case LINEERR_DIALDIALTONE '(= &H80000009
            msg = "Dial Dialtone"
        Case LINEERR_DIALPROMPT '(= &H8000000A
            msg = "Dial Prompt"
        Case LINEERR_DIALQUIET '(= &H8000000B
            msg = "Dial Quiet"
        Case LINEERR_INCOMPATIBLEAPIVERSION '(= &H8000000C
            msg = "Incompatible API Version"
        Case LINEERR_INCOMPATIBLEEXTVERSION '(= &H8000000D
            msg = "Incompatible Ext Version"
        Case LINEERR_INIFILECORRUPT '(= &H8000000E
            msg = "Ini File Corrupt"
        Case LINEERR_INUSE '(= &H8000000F
            msg = "In Use"
        Case LINEERR_INVALADDRESS '(= &H80000010
            msg = "Invalid Address"
        Case LINEERR_INVALADDRESSID '(= &H80000011
            msg = "Invalid Address ID"
        Case LINEERR_INVALADDRESSMODE '(= &H80000012
            msg = "Invalid Address Mode"
        Case LINEERR_INVALADDRESSSTATE '(= &H80000013
            msg = "Invalid Address State"
        Case LINEERR_INVALAPPHANDLE '(= &H80000014
            msg = "Invalid App Handle"
        Case LINEERR_INVALAPPNAME '(= &H80000015
            msg = "Invalid App Name"
        Case LINEERR_INVALBEARERMODE '(= &H80000016
            msg = "Invalid Bearer Mode"
        Case LINEERR_INVALCALLCOMPLMODE '(= &H80000017
            msg = "Invalid Call Completion Mode"
        Case LINEERR_INVALCALLHANDLE '(= &H80000018
            msg = "Invalid Call Handle"
        Case LINEERR_INVALCALLPARAMS '(= &H80000019
            msg = "Invalid Call Params"
        Case LINEERR_INVALCALLPRIVILEGE '(= &H8000001A
            msg = "Invalid Call Privilege"
        Case LINEERR_INVALCALLSELECT '(= &H8000001B
            msg = "Invalid Call Select"
        Case LINEERR_INVALCALLSTATE '(= &H8000001C
            msg = "Invalid Call State"
        Case LINEERR_INVALCALLSTATELIST '(= &H8000001D
            msg = "Invalid Call State List"
        Case LINEERR_INVALCARD '(= &H8000001E
            msg = "Invalid Card"
        Case LINEERR_INVALCOMPLETIONID '(= &H8000001F
            msg = "Invalid Completion ID"
        Case LINEERR_INVALCONFCALLHANDLE '(= &H80000020
            msg = "Invalid Conf Call Handle"
        Case LINEERR_INVALCONSULTCALLHANDLE '(= &H80000021
            msg = "Invalid Consult Call Handle"
        Case LINEERR_INVALCOUNTRYCODE '(= &H80000022
            msg = "Invalid Country Code"
        Case LINEERR_INVALDEVICECLASS '(= &H80000023
            msg = "Invalid Device Class"
        Case LINEERR_INVALDEVICEHANDLE '(= &H80000024
            msg = "Invalid Device Handle"
        Case LINEERR_INVALDIGITLIST '(= &H80000026
            msg = "Invalid Digit List"
        Case LINEERR_INVALDIGITMODE '(= &H80000027
            msg = "Invalid Digit Mode"
        Case LINEERR_INVALDIGITS '(= &H80000028
            msg = "Invalid Digits"
        Case LINEERR_INVALEXTVERSION '(= &H80000029
            msg = "Invalid Ext Version"
        Case LINEERR_INVALGROUPID '(= &H8000002A
            msg = "Invalid Group ID"
        Case LINEERR_INVALLINEHANDLE '(= &H8000002B
            msg = "Invalid Line Handle"
        Case LINEERR_INVALLINESTATE '(= &H8000002C
            msg = "Invalid Line State"
        Case LINEERR_INVALLOCATION '(= &H8000002D
            msg = "Invalid Location"
        Case LINEERR_INVALMEDIALIST '(= &H8000002E
            msg = "Invalid Media List"
        Case LINEERR_INVALMEDIAMODE '(= &H8000002F
            msg = "Invalid Media Mode"
        Case LINEERR_INVALMESSAGEID '(= &H80000030
            msg = "Invalid Message ID"
        Case LINEERR_INVALPARAM '(= &H80000032
            msg = "Invalid Param"
        Case LINEERR_INVALPARKID '(= &H80000033
            msg = "Invalid Park ID"
        Case LINEERR_INVALPARKMODE '(= &H80000034
            msg = "Invalid Park Mode"
        Case LINEERR_INVALPOINTER '(= &H80000035
            msg = "Invalid Pointer"
        Case LINEERR_INVALPRIVSELECT '(= &H80000036
            msg = "Invalid Priv Select"
        Case LINEERR_INVALRATE '(= &H80000037
            msg = "Invalid Rate"
        Case LINEERR_INVALREQUESTMODE '(= &H80000038
            msg = "Invalid Request Mode"
        Case LINEERR_INVALTERMINALID '(= &H80000039
            msg = "Invalid Terminal ID"
        Case LINEERR_INVALTERMINALMODE '(= &H8000003A
            msg = "Invalid Terminal Mode"
        Case LINEERR_INVALTIMEOUT '(= &H8000003B
            msg = "Invalid Time Out"
        Case LINEERR_INVALTONE '(= &H8000003C
            msg = "Invalid Tone"
        Case LINEERR_INVALTONELIST '(= &H8000003D
            msg = "Invalid Tone List"
        Case LINEERR_INVALTONEMODE '(= &H8000003E
            msg = "Invalid Tone Mode"
        Case LINEERR_INVALTRANSFERMODE '(= &H8000003F
            msg = "Invalid Transfer Mode"
        Case LINEERR_LINEMAPPERFAILED '(= &H80000040
            msg = "Line Mapper Failed"
        Case LINEERR_NOCONFERENCE '(= &H80000041
            msg = "No Conference"
        Case LINEERR_NODEVICE '(= &H80000042
            msg = "No Device"
        Case LINEERR_NODRIVER '(= &H80000043
            msg = "No Driver"
        Case LINEERR_NOMEM '(= &H80000044
            msg = "No Memory"
        Case LINEERR_NOREQUEST '(= &H80000045
            msg = "No Request"
        Case LINEERR_NOTOWNER '(= &H80000046
            msg = "Not Owner"
        Case LINEERR_NOTREGISTERED '(= &H80000047
            msg = "Not Registered"
        Case LINEERR_OPERATIONFAILED '(= &H80000048
            msg = "Operation Failed"
        Case LINEERR_OPERATIONUNAVAIL '(= &H80000049
            msg = "Operation Unavailable"
        Case LINEERR_RATEUNAVAIL '(= &H8000004A
            msg = "Rate Unavailable"
        Case LINEERR_RESOURCEUNAVAIL '(= &H8000004B
            msg = "Resource Unavailable"
        Case LINEERR_REQUESTOVERRUN '(= &H8000004C
            msg = "Request Overrun"
        Case LINEERR_STRUCTURETOOSMALL '(= &H8000004D
            msg = "Structure Too Small"
        Case LINEERR_TARGETNOTFOUND '(= &H8000004E
            msg = "Target Not found"
        Case LINEERR_TARGETSELF '(= &H8000004F
            msg = "Target Self"
        Case LINEERR_UNINITIALIZED '(= &H80000050
            msg = "Uninitialized"
        Case LINEERR_USERUSERINFOTOOBIG '(= &H80000051
            msg = "UserUser Info Too Big"
        Case LINEERR_REINIT '(= &H80000052
            msg = "Re-init"
        Case LINEERR_ADDRESSBLOCKED '(= &H80000053
            msg = "Address Blocked"
        Case LINEERR_BILLINGREJECTED '(= &H80000054
            msg = "Billing Rejected"
        Case LINEERR_INVALFEATURE '(= &H80000055
            msg = "Invalid Feature"
        Case LINEERR_NOMULTIPLEINSTANCE '(= &H80000056
            msg = "No Multiple Instance"
        Case Else
            msg = "Unknown Error" ' undefined
    End Select
    
    GetLineErrString = msg
End Function

Public Function GetLineStateString(ByVal state As Long) As String
    Dim msg As String

    Select Case state
        Case LINECALLSTATE_IDLE                       '&H1
            msg = "idle"
        
        Case LINECALLSTATE_OFFERING                   '&H2
            msg = "offering call"
        
        Case LINECALLSTATE_ACCEPTED                   '&H4
            msg = "accepted"
        
        Case LINECALLSTATE_DIALTONE                   '&H8
            msg = "dial-tone detected"
        
        Case LINECALLSTATE_DIALING                    '&H10
            msg = "dialing"
        
        Case LINECALLSTATE_RINGBACK                   '&H20
            msg = "ring-back detected"
        
        Case LINECALLSTATE_BUSY                       '&H40
            msg = "busy detected"
        
        Case LINECALLSTATE_SPECIALINFO                '&H80
            msg = "network error"
        
        Case LINECALLSTATE_CONNECTED                  '&H100
            msg = "connected"
        
        Case LINECALLSTATE_PROCEEDING                 '&H200
            msg = "proceeding"
        
        Case LINECALLSTATE_ONHOLD                     '&H400
            msg = "on hold"
        
        Case LINECALLSTATE_CONFERENCED                '&H800
            msg = "connected to conference"
        
        Case LINECALLSTATE_ONHOLDPENDCONF             '&H1000
            msg = "connecting to conference"
        
        Case LINECALLSTATE_ONHOLDPENDTRANSFER         '&H2000
            msg = "transferring"
        
        Case LINECALLSTATE_DISCONNECTED               '&H4000
            msg = "disconnected"
        
        Case LINECALLSTATE_UNKNOWN                    '&H8000
            msg = "unknown call state"
        
        Case Else
            msg = "unknown value passed to GetLineStateString()"
        
    End Select
    
    GetLineStateString = msg
End Function
