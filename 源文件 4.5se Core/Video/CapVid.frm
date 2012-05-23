VERSION 5.00
Begin VB.Form frmCapVid 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "捕获视频流"
   ClientHeight    =   2760
   ClientLeft      =   345
   ClientTop       =   1545
   ClientWidth     =   5745
   Icon            =   "CapVid.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "捕获:"
      Height          =   1725
      Left            =   225
      TabIndex        =   5
      Top             =   225
      Width           =   3525
      Begin VB.TextBox txtFps 
         Height          =   285
         Left            =   2115
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "15"
         Top             =   270
         Width           =   690
      End
      Begin VB.CheckBox chkLimit 
         Caption         =   "使用时间限制:"
         Height          =   285
         Left            =   315
         TabIndex        =   8
         Top             =   765
         Width           =   1680
      End
      Begin VB.TextBox txtSec 
         Height          =   285
         Left            =   2115
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "30"
         Top             =   765
         Width           =   690
      End
      Begin VB.CheckBox chkAudio 
         Caption         =   "捕获音频"
         Height          =   285
         Left            =   315
         TabIndex        =   6
         Top             =   1260
         Width           =   1905
      End
      Begin VB.Label lblStaticText 
         Caption         =   "帧率 (帧/秒):"
         Height          =   240
         Index           =   0
         Left            =   315
         TabIndex        =   11
         Top             =   315
         Width           =   1545
      End
      Begin VB.Label lblStaticText 
         Caption         =   "秒"
         Height          =   285
         Index           =   1
         Left            =   3015
         TabIndex        =   10
         Top             =   810
         Width           =   330
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "压缩(&P)..."
      Height          =   330
      Index           =   4
      Left            =   4050
      TabIndex        =   4
      Top             =   1440
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "视频(&V)..."
      Height          =   330
      Index           =   3
      Left            =   4050
      TabIndex        =   3
      Top             =   360
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "音频(&A)..."
      Height          =   330
      Index           =   2
      Left            =   4050
      TabIndex        =   2
      Top             =   900
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消(&C)"
      Height          =   330
      Index           =   1
      Left            =   4140
      TabIndex        =   1
      Top             =   2160
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&Y)"
      Height          =   330
      Index           =   0
      Left            =   2655
      TabIndex        =   0
      Top             =   2160
      Width           =   1185
   End
End
Attribute VB_Name = "frmCapVid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CapParams As CAPTUREPARMS

Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0 'OK
        Call ProcessCapInfo
        Me.Hide
    Case 1 'Cancel
        Unload Me
    Case 2 'Audio
        Call SetAudioFormatDlg(Me.hWnd)
    Case 3 'Video
        Call capDlgVideoFormat(frmMain.capwnd)
        Call ResizeCaptureWindow(frmMain.capwnd)
    Case 4 'Compress
        Call capDlgVideoCompression(frmMain.capwnd)
End Select
End Sub


Private Sub ProcessCapInfo()


With CapParams
'   // set the defaults we won't bother the user with
'   show message after pre-roll
    .fMakeUserHitOKToCapture = -(True) ' - converts VB Boolean to C BOOL
'   in case we use error callbacks later
    .wPercentDropForError = 10
'   fUsingDOSMemory is obsolete
    .fUsingDOSMemory = False
'   The number of video buffers should be enough to get through
'   disk seeks and thermal recalibrations
    .wNumVideoRequested = 32
'   Do abort on the left mouse
    .fAbortLeftMouse = -(True)
'   Do abort on the right mouse
    .fAbortRightMouse = -(True) '- converts VB boolean to C BOOL
'   If wChunkGranularity is zero, the granularity will be set to the
'   disk sector size.
    .wChunkGranularity = 0
'   use default
    .dwAudioBufferSize = 0
'   attempt to disable caching
    .fDisableWriteCache = -(True)
'   not using MCI
    .fMCIControl = False
    .fStepCaptureAt2x = False
'   not multi-threading
    .fYield = False
'   request audio buffers
    .wNumAudioRequested = 4 '10 is max limit

'   //these parameters are loaded from registry
    If "音频" = Trim$(UCase(GetSetting(App.Title, "选项", "流主管", "音频"))) Then
        .AVStreamMaster = AVSTREAMMASTER_AUDIO 'use audio clock to synchronize AVI
    Else
        .AVStreamMaster = AVSTREAMMASTER_NONE
    End If
    'set index size of AVI file (max frames)
    .dwIndexSize = Val(GetSetting(App.Title, "选项", "最大帧数", INDEX_15_MINUTES))

'   //Now set the parameters from the UserInput
    .dwRequestMicroSecPerFrame = microsSecFromFPS(Val(txtFps.Text))
    .fCaptureAudio = -(CBool(chkAudio.Value))
    .fLimitEnabled = -(CBool(chkLimit.Value))
    .wTimeLimit = Val(txtSec.Text)

End With
'set the new setup info
Call capCaptureSetSetup(frmMain.capwnd, CapParams)
'Kludgy - but...
Me.Tag = True 'this tells main form that OK button was pushed
End Sub

Private Function microsSecFromFPS(ByVal fps As Long) As Long
'note I am not using floating point here so these are not too exact
If fps = 0 Then Exit Function 'avoid divide by 0 errors
microsSecFromFPS = 1000000 / fps
End Function

Private Sub txtFps_LostFocus()
If Val(txtFps.Text) < 1 Then txtFps.Text = "1"
If Val(txtFps.Text) > 100 Then txtFps.Text = "100"
End Sub
Private Sub txtFPS_KeyPress(KeyAscii As Integer)
'allow backspace key
If KeyAscii = 8 Then Exit Sub
'logic to keep the user input valid
If KeyAscii < 48 Then KeyAscii = 0
If KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtSec_KeyPress(KeyAscii As Integer)
'allow backspace key
If KeyAscii = 8 Then Exit Sub
'logic to keep the user input valid
If KeyAscii < 48 Then KeyAscii = 0
If KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub Form_Load()
'this form loads settings automatically each time it is loaded
Call LoadMe
End Sub

Private Sub Form_Unload(Cancel As Integer)
'this form saves settings automatically each time it is unloaded
Call SaveMe
End Sub

Private Sub LoadMe()
    txtFps.Text = GetSetting(App.Title, "vidcap settings", "fps", "15")
    chkLimit.Value = Val(GetSetting(App.Title, "vidcap settings", "time limit", "0"))
    txtSec.Text = GetSetting(App.Title, "vidcap settings", "seconds", "30")
    chkAudio.Value = Val(GetSetting(App.Title, "vidcap settings", "cap audio", "0"))
    
End Sub

Private Sub SaveMe()
    Call SaveSetting(App.Title, "vidcap settings", "fps", txtFps.Text)
    Call SaveSetting(App.Title, "vidcap settings", "time limit", chkLimit.Value)
    Call SaveSetting(App.Title, "vidcap settings", "seconds", txtSec.Text)
    Call SaveSetting(App.Title, "vidcap settings", "cap audio", chkAudio.Value)
End Sub


