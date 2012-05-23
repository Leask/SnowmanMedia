VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form10 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form10"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5325
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MCI.MMControl MMControl1 
      Height          =   300
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   529
      _Version        =   393216
      BorderStyle     =   0
      PrevEnabled     =   -1  'True
      NextEnabled     =   -1  'True
      PlayEnabled     =   -1  'True
      PauseEnabled    =   -1  'True
      BackEnabled     =   -1  'True
      StepEnabled     =   -1  'True
      StopEnabled     =   -1  'True
      EjectEnabled    =   -1  'True
      Silent          =   -1  'True
      RecordVisible   =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3150
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打开VCD文件"
      Height          =   285
      Left            =   3690
      TabIndex        =   0
      Top             =   270
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VCD"
      Height          =   240
      Left            =   0
      TabIndex        =   5
      Top             =   315
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3645
      Left            =   0
      Picture         =   "Form10.frx":0000
      ScaleHeight     =   3615
      ScaleWidth      =   5280
      TabIndex        =   4
      Top             =   585
      Width           =   5310
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   645
      Left            =   1395
      TabIndex        =   3
      Top             =   1260
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "媒体信息"
      Height          =   180
      Left            =   495
      TabIndex        =   2
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   315
      TabIndex        =   1
      Top             =   1620
      Width           =   90
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim intpos As Integer
Dim ingreturnvalue As Long
Const HWND_TOPMOST = -1
Const SWP_SHOWWINDOW = &H40

Const INTERVAL = 50
Const INTERVAL_PLUS = 55
Dim c As String
Dim b As Integer
Dim a As String


Dim CurVal As Double
Private Sub command2_Click()
    MMControl1.Command = "Stop"
    MMControl1.Command = "Close"
    
    Form10.CommonDialog1.Filter = "VCD文件(*.dat)" & _
    "|*.dat|浏览所有文件:*.*|*.*"
    
    Form10.CommonDialog1.FilterIndex = 1
    Form10.CommonDialog1.ShowOpen
    If Len(Form10.CommonDialog1.filename) > 0 Then
    
    MMControl1.filename = Form10.CommonDialog1.filename
    Form10.Caption = Form10.CommonDialog1.filename & " - Snowman Media  CD/VCD Player"
  MMControl1.DeviceType = ""
  MMControl1.Command = "Open"
  MMControl1.Command = "PLay"
End If
End Sub

Private Sub Command1_Click()

If Form10.Command1.Caption = "VCD" Then
'MMControl1.Command = "Close"
Form10.Height = 4515
Form10.Width = 5370
Form10.Command1.Caption = "CD"
'MMControl1.DeviceType = ""
'MMControl1.Command = "Open"

Else
'MMControl1.Command = "Close"
Form10.Height = 870
Form10.Width = 3225
Form10.Command1.Caption = "VCD"
'MMControl1.DeviceType = "CDaudio" 'MCI设备类型为CD唱片
'MMControl1.Command = "Open"
End If
End Sub

Private Sub Form_Load()
 ingreturnvalue = SetWindowPos(Me.hwnd, HWND_TOPMOST, Val(10), Val(10), Val(10), Val(10), SWP_SHOWWINDOW)
    Form10.Height = 870
Form10.Width = 3225
Form10.Command1.Caption = "VCD"

MMControl1.DeviceType = "CDaudio" 'MCI设备类型为CD唱片
'MMControl1.filename = "E:\My Documents\My Videos\ge.dat"
      MMControl1.Command = "Open" '打开设备
MMControl1.TimeFormat = 0
MMControl1.UpdateInterval = 1000
Form10.Caption = "Snowman Media  CD/VCD Player"
CurVal = 0#
HScroll1.value = 0
End Sub
 Private Sub Form_Unload(Cancel As Integer)

        MMControl1.Command = "Close" '退出时关闭MCI设备

        End Sub


Private Sub MMControl1_EjectClick(Cancel As Integer)
''c = MMControl1.DeviceType

''MMControl1.DeviceType = "CDaudio"
'MMControl1.Command = "Eject"
Form10.Caption = "正在处理光驱 - Snowman Media  CD/VCD Player"


End Sub

Private Sub MMControl1_EjectCompleted(Errorcode As Long)

''MMControl1.DeviceType = c

Form10.Caption = "已经完成处理 - Snowman Media  CD/VCD Player"

End Sub

Private Sub MMControl1_NextClick(Cancel As Integer)
Label2.Caption = "正在播放第" + Str$(MMControl1.Track) + "首曲目"
a = "媒体共有"
a = a + Trim(Str$(MMControl1.Tracks))
a = a + "首曲目,总时间"
a = a + Trim(Str$((MMControl1.Length / 1000) \ 60))
Label1.Caption = a + "分钟"
Form10.Caption = Label2.Caption + " - Snowman Media  CD/VCD Player"
End Sub

Private Sub MMControl1_PauseClick(Cancel As Integer)
Form10.Caption = "已经暂停 - Snowman Media  CD Player"
MMControl1.UpdateInterval = 0
End Sub




Private Sub MMControl1_PlayClick(Cancel As Integer)
Form10.Caption = Label2.Caption + " - Snowman Media  CD/VCD Player"
MMControl1.UpdateInterval = INTERVAL
End Sub





Private Sub MMControl1_PrevClick(Cancel As Integer)
Label2.Caption = "正在播放第" + Str$(MMControl1.Track) + "首曲目"
a = "媒体共有"
a = a + Trim(Str$(MMControl1.Tracks))
a = a + "首曲目,总时间"
a = a + Trim(Str$((MMControl1.Length / 1000) \ 60))
Label1.Caption = a + "分钟"
Form10.Caption = Label2.Caption + " - Snowman Media  CD/VCD Player"
HScroll1.value = 0
CurVal = 0#
End Sub

Private Sub mmcontrol1_StatusUpdate()
MMControl1.hWndDisplay = Picture1.hwnd
Label2.Caption = "正在播放第" + Str$(MMControl1.Track) + "首曲目"
a = "媒体共有"
a = a + Trim(Str$(MMControl1.Tracks))
a = a + "首曲目,总时间"
a = a + Trim(Str$((MMControl1.Length / 1000) \ 60))
Label1.Caption = a + "分钟"

Const mci_mode_play = 526
Dim value As Integer
If Not MMControl1.Mode = mci_mode_play Then
HScroll1.value = HScroll1.Max
MMControl1.UpdateInterval = 0
Exit Sub
End If
CurVal = CurVal + INTERVAL_PLUS
value = CInt(CurVal / MMControl1.Length) * 100
If value > HScroll1.Max Then
value = 100
End If
HScroll1.value = value
End Sub
Private Sub MMControl1_StopClick(Cancel As Integer)
Form10.Caption = "已经停止 - Snowman Media  CD/VCD Player"
End Sub

