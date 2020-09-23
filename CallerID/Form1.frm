VERSION 5.00
Object = "{49FE71CE-3D83-45E4-B057-E42CA42F462B}#15.0#0"; "srCallerID.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caller ID ocx test"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Begin srCallerID.TAPICallerID TAPICallerID1 
      Left            =   510
      Top             =   1095
      _ExtentX        =   1138
      _ExtentY        =   1138
      LineID          =   13
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   165
      TabIndex        =   7
      Top             =   1875
      Width           =   3705
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Last Call Info"
      Height          =   390
      Left            =   1530
      TabIndex        =   5
      Top             =   1260
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   390
      Left            =   2835
      TabIndex        =   0
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Debug"
      Height          =   210
      Left            =   165
      TabIndex        =   6
      Top             =   1650
      Width           =   540
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1305
      TabIndex        =   4
      Top             =   750
      Width           =   2490
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Caller Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   780
      Width           =   990
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1305
      TabIndex        =   2
      Top             =   255
      Width           =   2490
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Caller ID:"
      Height          =   315
      Left            =   495
      TabIndex        =   1
      Top             =   300
      Width           =   795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'* Module:  Form1.frm
'* Purpose: Demonstrate using the the srCallerID.ocx
'*
'* Author:  Sumanta Ray
'* Email:   chandsumant@hotmail.com
'* Date:    14/07/2005
'* Copyright (c) 2003-2005 Sumanta Ray. All rights reserved
'***********************************************************************

Private Sub Command1_Click()
    Label2 = ""
    Label4 = ""
End Sub

Private Sub Command2_Click()
    Label2 = TAPICallerID1.LastCallerID
    Label4 = TAPICallerID1.LastCallerName
End Sub

Private Sub Form_Load()
    TAPICallerID1.StartMonitor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TAPICallerID1.StopMonitor
End Sub

Private Sub TAPICallerID1_OnCallerID(ByVal CID As String)
    Label2 = CID
End Sub

Private Sub TAPICallerID1_OnCallerName(ByVal CNAME As String)
    Label4 = CNAME
End Sub

Private Sub TAPICallerID1_OnDebug(ByVal msg As String)
    List1.AddItem msg
    List1.ListIndex = List1.ListCount - 1
End Sub

Private Sub TAPICallerID1_OnError(ByVal msg As String, ByVal source As String)
    List1.AddItem msg & " - " & source
    List1.ListIndex = List1.ListCount - 1
End Sub

