VERSION 5.00
Object = "{21D4D402-6A96-11D2-A0F9-444553540000}#1.0#0"; "EQPRO.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "EQ  - Snowman Media  2.0"
   ClientHeight    =   1065
   ClientLeft      =   4710
   ClientTop       =   4995
   ClientWidth     =   4185
   Icon            =   "EQPro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMonitorMutes 
      Interval        =   250
      Left            =   1815
      Top             =   2430
   End
   Begin VB.Frame Frame1 
      Caption         =   "项目:"
      Height          =   1065
      Left            =   675
      TabIndex        =   2
      Top             =   0
      Width           =   3510
      Begin EQPro.ucEQPro eqCustom 
         Height          =   450
         Left            =   45
         TabIndex        =   5
         Top             =   540
         Width           =   3390
         _ExtentX        =   794
         _ExtentY        =   794
         LayoutAlign     =   0
      End
      Begin VB.CheckBox chkMuteCustom 
         Caption         =   "CD Audio"
         Height          =   195
         Left            =   1845
         TabIndex        =   4
         Top             =   315
         Width           =   1545
      End
      Begin VB.ComboBox cmbCustom 
         Height          =   300
         ItemData        =   "EQPro.frx":1582
         Left            =   45
         List            =   "EQPro.frx":1592
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   225
         Width           =   1770
      End
   End
   Begin VB.CheckBox chkMuteMaster 
      Caption         =   "静音"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   855
      Width           =   825
   End
   Begin EQPro.ucEQPro eqMaster 
      Height          =   690
      Left            =   90
      TabIndex        =   6
      Top             =   180
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "主音量"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkMuteCustom_Click()
    eqCustom.Mute = -chkMuteCustom.Value
End Sub
Private Sub chkMuteMaster_Click()
    eqMaster.Mute = -chkMuteMaster.Value
End Sub
Private Sub cmbCustom_Click()
    eqCustom.VolControl = cmbCustom.ListIndex + 1
    chkMuteCustom.Caption = "Mute " + cmbCustom.List(cmbCustom.ListIndex)
End Sub
Private Sub Form_Load()
    cmbCustom.ListIndex = 0
End Sub
Private Sub tmrMonitorMutes_Timer()
    chkMuteMaster.Value = Abs(eqMaster.Mute)
    chkMuteCustom.Value = Abs(eqCustom.Mute)
End Sub
