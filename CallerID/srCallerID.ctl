VERSION 5.00
Begin VB.UserControl TAPICallerID 
   BackColor       =   &H00D9B59F&
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1860
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "srCallerID.ctx":0000
   ScaleHeight     =   1620
   ScaleWidth      =   1860
   ToolboxBitmap   =   "srCallerID.ctx":0018
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "2317.."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   210
      TabIndex        =   0
      Top             =   480
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   90
      Picture         =   "srCallerID.ctx":032A
      Stretch         =   -1  'True
      Top             =   45
      Width           =   450
   End
End
Attribute VB_Name = "TAPICallerID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***********************************************************************
'* Module:  srTapi.ctl
'* Purpose: UserControl, exposes the TAPI interface to user
'*
'* Author:  Sumanta Ray
'* Email:   chandsumant@hotmail.com
'* Date:    14/07/2005
'* Copyright (c) 2003-2005 Sumanta Ray. All rights reserved
'***********************************************************************

Option Explicit
'all enums
'***************

Enum smTAPILineErrorConstants
    LINEERR_ALLOCATED = &H80000001
    LINEERR_BADDEVICEID = &H80000002
    LINEERR_BEARERMODEUNAVAIL = &H80000003
    LINEERR_CALLUNAVAIL = &H80000005
    LINEERR_COMPLETIONOVERRUN = &H80000006
    LINEERR_CONFERENCEFULL = &H80000007
    LINEERR_DIALBILLING = &H80000008
    LINEERR_DIALDIALTONE = &H80000009
    LINEERR_DIALPROMPT = &H8000000A
    LINEERR_DIALQUIET = &H8000000B
    LINEERR_INCOMPATIBLEAPIVERSION = &H8000000C
    LINEERR_INCOMPATIBLEEXTVERSION = &H8000000D
    LINEERR_INIFILECORRUPT = &H8000000E
    LINEERR_INUSE = &H8000000F
    LINEERR_INVALADDRESS = &H80000010
    LINEERR_INVALADDRESSID = &H80000011
    LINEERR_INVALADDRESSMODE = &H80000012
    LINEERR_INVALADDRESSSTATE = &H80000013
    LINEERR_INVALAPPHANDLE = &H80000014
    LINEERR_INVALAPPNAME = &H80000015
    LINEERR_INVALBEARERMODE = &H80000016
    LINEERR_INVALCALLCOMPLMODE = &H80000017
    LINEERR_INVALCALLHANDLE = &H80000018
    LINEERR_INVALCALLPARAMS = &H80000019
    LINEERR_INVALCALLPRIVILEGE = &H8000001A
    LINEERR_INVALCALLSELECT = &H8000001B
    LINEERR_INVALCALLSTATE = &H8000001C
    LINEERR_INVALCALLSTATELIST = &H8000001D
    LINEERR_INVALCARD = &H8000001E
    LINEERR_INVALCOMPLETIONID = &H8000001F
    LINEERR_INVALCONFCALLHANDLE = &H80000020
    LINEERR_INVALCONSULTCALLHANDLE = &H80000021
    LINEERR_INVALCOUNTRYCODE = &H80000022
    LINEERR_INVALDEVICECLASS = &H80000023
    LINEERR_INVALDEVICEHANDLE = &H80000024
    LINEERR_INVALDIGITLIST = &H80000026
    LINEERR_INVALDIGITMODE = &H80000027
    LINEERR_INVALDIGITS = &H80000028
    LINEERR_INVALEXTVERSION = &H80000029
    LINEERR_INVALGROUPID = &H8000002A
    LINEERR_INVALLINEHANDLE = &H8000002B
    LINEERR_INVALLINESTATE = &H8000002C
    LINEERR_INVALLOCATION = &H8000002D
    LINEERR_INVALMEDIALIST = &H8000002E
    LINEERR_INVALMEDIAMODE = &H8000002F
    LINEERR_INVALMESSAGEID = &H80000030
    LINEERR_INVALPARAM = &H80000032
    LINEERR_INVALPARKID = &H80000033
    LINEERR_INVALPARKMODE = &H80000034
    LINEERR_INVALPOINTER = &H80000035
    LINEERR_INVALPRIVSELECT = &H80000036
    LINEERR_INVALRATE = &H80000037
    LINEERR_INVALREQUESTMODE = &H80000038
    LINEERR_INVALTERMINALID = &H80000039
    LINEERR_INVALTERMINALMODE = &H8000003A
    LINEERR_INVALTIMEOUT = &H8000003B
    LINEERR_INVALTONE = &H8000003C
    LINEERR_INVALTONELIST = &H8000003D
    LINEERR_INVALTONEMODE = &H8000003E
    LINEERR_INVALTRANSFERMODE = &H8000003F
    LINEERR_LINEMAPPERFAILED = &H80000040
    LINEERR_NOCONFERENCE = &H80000041
    LINEERR_NODEVICE = &H80000042
    LINEERR_NODRIVER = &H80000043
    LINEERR_NOMEM = &H80000044
    LINEERR_NOREQUEST = &H80000045
    LINEERR_NOTOWNER = &H80000046
    LINEERR_NOTREGISTERED = &H80000047
    LINEERR_OPERATIONFAILED = &H80000048
    LINEERR_OPERATIONUNAVAIL = &H80000049
    LINEERR_RATEUNAVAIL = &H8000004A
    LINEERR_RESOURCEUNAVAIL = &H8000004B
    LINEERR_REQUESTOVERRUN = &H8000004C
    LINEERR_STRUCTURETOOSMALL = &H8000004D
    LINEERR_TARGETNOTFOUND = &H8000004E
    LINEERR_TARGETSELF = &H8000004F
    LINEERR_UNINITIALIZED = &H80000050
    LINEERR_USERUSERINFOTOOBIG = &H80000051
    LINEERR_REINIT = &H80000052
    LINEERR_ADDRESSBLOCKED = &H80000053
    LINEERR_BILLINGREJECTED = &H80000054
    LINEERR_INVALFEATURE = &H80000055
    LINEERR_NOMULTIPLEINSTANCE = &H80000056
End Enum

Enum smTAPILinePrivileges
    LINECALLPRIVILEGE_NONE = &H1&
    LINECALLPRIVILEGE_MONITOR = &H2&
    LINECALLPRIVILEGE_OWNER = &H4&
End Enum

Enum smTAPILineCallStates
    LINECALLSTATE_IDLE = &H1&
    LINECALLSTATE_OFFERING = &H2&
    LINECALLSTATE_ACCEPTED = &H4&
    LINECALLSTATE_DIALTONE = &H8&
    LINECALLSTATE_DIALING = &H10&
    LINECALLSTATE_RINGBACK = &H20&
    LINECALLSTATE_BUSY = &H40&
    LINECALLSTATE_SPECIALINFO = &H80&
    LINECALLSTATE_CONNECTED = &H100&
    LINECALLSTATE_PROCEEDING = &H200&
    LINECALLSTATE_ONHOLD = &H400&
    LINECALLSTATE_CONFERENCED = &H800&
    LINECALLSTATE_ONHOLDPENDCONF = &H1000&
    LINECALLSTATE_ONHOLDPENDTRANSFER = &H2000&
    LINECALLSTATE_DISCONNECTED = &H4000&
    LINECALLSTATE_UNKNOWN = &H8000&
End Enum

'LINECALLPARTYID_ Constants
Enum smTAPILineCallPartyConstants
    LINECALLPARTYID_BLOCKED = &H1
    LINECALLPARTYID_OUTOFAREA = &H2
    LINECALLPARTYID_NAME = &H4
    LINECALLPARTYID_ADDRESS = &H8
    LINECALLPARTYID_PARTIAL = &H10
    LINECALLPARTYID_UNKNOWN = &H20
    LINECALLPARTYID_UNAVAIL = &H40
End Enum

Enum smTAPIAddressModes
    LINEADDRESSMODE_ADDRESSID = &H1&
    LINEADDRESSMODE_DIALABLEADDR = &H2&
End Enum

Enum smTAPIBearerModes
    LINEBEARERMODE_VOICE = &H1&
    LINEBEARERMODE_SPEECH = &H2&
    LINEBEARERMODE_MULTIUSE = &H4&
    LINEBEARERMODE_DATA = &H8&
    LINEBEARERMODE_ALTSPEECHDATA = &H10&
    LINEBEARERMODE_NONCALLSIGNALING = &H20&
End Enum

Enum smTAPIMediaModes
    LINEMEDIAMODE_UNKNOWN = &H2&
    LINEMEDIAMODE_INTERACTIVEVOICE = &H4&
    LINEMEDIAMODE_AUTOMATEDVOICE = &H8&
    LINEMEDIAMODE_DATAMODEM = &H10&
    LINEMEDIAMODE_G3FAX = &H20&
    LINEMEDIAMODE_TDD = &H40&
    LINEMEDIAMODE_G4FAX = &H80&
    LINEMEDIAMODE_DIGITALDATA = &H100&
    LINEMEDIAMODE_TELETEX = &H200&
    LINEMEDIAMODE_VIDEOTEX = &H400&
    LINEMEDIAMODE_TELEX = &H800&
    LINEMEDIAMODE_MIXED = &H1000&
    LINEMEDIAMODE_ADSI = &H2000&
End Enum
'*****************

'Events
Event OnCallerID(ByVal CID As String)
Event OnCallerName(ByVal CNAME As String)
Event OnError(ByVal msg As String, ByVal source As String)
Event OnDebug(ByVal msg As String)

Private WithEvents objTapiLine As TAPILINE
Attribute objTapiLine.VB_VarHelpID = -1

Private Sub objTapiLine_OnCallerID(ByVal CID As String)
    RaiseEvent OnCallerID(CID)
End Sub

Private Sub objTapiLine_OnCallerName(ByVal CNAME As String)
    RaiseEvent OnCallerName(CNAME)
End Sub

Private Sub objTapiLine_OnDebug(ByVal msg As String)
    RaiseEvent OnDebug(msg)
End Sub

Private Sub objTapiLine_OnError(ByVal msg As String, ByVal source As String)
    RaiseEvent OnError(msg, source)
End Sub

Private Sub UserControl_Initialize()
    Height = 650
    Width = 650
    Set objTapiLine = New TAPILINE
    'initialize all the lines
    '****************************
    Dim success As Boolean
    success = objTapiLine.InitLines
    If Not success Then
        RaiseEvent OnError("Failed to initialize the TAPI lines", "InitLines")
    End If
    '****************************
End Sub

Public Sub StartMonitor()
    'trying to open the TAPI line
    '****************************
    If objTapiLine.LineSupportsVoiceCalls = True Then 'if voice call supported by the line device
        If objTapiLine.OpenLine(LINECALLPRIVILEGE_MONITOR Or LINECALLPRIVILEGE_OWNER, LINEMEDIAMODE_INTERACTIVEVOICE) = False Then
            RaiseEvent OnError("Failed to open TAPI line", "OpenLine")
        End If
    Else 'if voice call not supported
        If objTapiLine.OpenLine(LINECALLPRIVILEGE_MONITOR Or LINECALLPRIVILEGE_OWNER, LINEMEDIAMODE_DATAMODEM) = False Then
            RaiseEvent OnError("Failed to open TAPI line", "OpenLine")
        End If
    End If
    '****************************
End Sub

Public Sub StopMonitor()
    objTapiLine.CloseLine
End Sub

Public Property Get LastCallerID() As String
    LastCallerID = objTapiLine.LastCallerID
End Property

Public Property Get LastCallerName() As String
    LastCallerName = objTapiLine.LastCallerName
End Property

Public Property Get LineName() As String
    LineName = objTapiLine.LineName
End Property

Public Property Get NumberOfLines() As Long
    NumberOfLines = objTapiLine.NumLines
End Property

Public Property Get LineID() As Long
    LineID = objTapiLine.CurrentLineID
End Property

Public Property Let LineID(ByVal LID As Long)
    objTapiLine.CurrentLineID = LID
    PropertyChanged "LineID"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

'Load property values from storage
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING LINES!
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    objTapiLine.CurrentLineID = PropBag.ReadProperty("LineID", 0)
End Sub

Private Sub UserControl_Terminate()
    Set objTapiLine = Nothing
End Sub

'Write property values to storage
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING LINES!
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    Call PropBag.WriteProperty("LineID", objTapiLine.CurrentLineID, 0)
End Sub

Public Property Get About() As Variant
Attribute About.VB_Description = "Shows the About Dialog"
Attribute About.VB_ProcData.VB_Invoke_Property = "PAbout"
'
End Property

Public Property Let About(ByVal vNewValue As Variant)
'
End Property

Public Function ErrorString(ByVal errCode As Long) As String
Attribute ErrorString.VB_Description = "Takes Error number as argument and shows the Error description"
    ErrorString = GetLineErrString(errCode)
End Function

Private Sub UserControl_Paint()
    DrawRaised
End Sub

Private Sub UserControl_Resize()
    Height = 650
    Width = 650
End Sub

Private Sub DrawRaised()
    Line (0, 0)-(Width, 0), vb3DDKShadow
    Line (0, 0)-(0, Height), vb3DDKShadow
    Line (Width - 5, 0)-(Width - 5, Height - 5), vb3DHighlight
    Line (0, Height - 5)-(Width - 5, Height - 5), vb3DHighlight
    
    Line (15, 15)-(ScaleWidth - 15, 15), vb3DHighlight
    Line (15, 15)-(15, ScaleHeight - 15), vb3DHighlight
    Line (ScaleWidth - 15, 15)-(ScaleWidth - 15, ScaleHeight - 15), vb3DDKShadow
    Line (15, ScaleHeight - 15)-(ScaleWidth - 15, ScaleHeight - 15), vb3DDKShadow
End Sub

