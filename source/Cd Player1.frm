VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form102 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Snowman Media  2.0"
   ClientHeight    =   5550
   ClientLeft      =   4140
   ClientTop       =   1980
   ClientWidth     =   6705
   ForeColor       =   &H00000000&
   Icon            =   "Cd Player1.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   5550
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Eject 
      Appearance      =   0  'Flat
      Caption         =   "Open"
      Enabled         =   0   'False
      Height          =   345
      Left            =   1890
      TabIndex        =   6
      ToolTipText     =   "Eject CD"
      Top             =   4905
      Width           =   1080
   End
   Begin VB.TextBox TimeWindow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   315
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "[00]00:00"
      ToolTipText     =   "播放时间  - Snowman Media  2.0"
      Top             =   2700
      Width           =   1455
   End
   Begin VB.VScrollBar Volume 
      Height          =   240
      Left            =   1575
      MouseIcon       =   "Cd Player1.frx":1582
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1620
      Width           =   150
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3060
      Top             =   4950
   End
   Begin VB.ComboBox TrackSelection 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   450
      MouseIcon       =   "Cd Player1.frx":16D4
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Text            =   "16"
      ToolTipText     =   "CD曲目选择  - Snowman Media  2.0"
      Top             =   1575
      Width           =   555
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   4095
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Audios Opening Window  - Snowman Media  2.0"
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      DragIcon        =   "Cd Player1.frx":1826
      Height          =   1050
      Left            =   3870
      TabIndex        =   5
      Top             =   4500
      Width           =   2985
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   65535
      DisplayForeColor=   16711680
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   -1  'True
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   -1  'True
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Image Image9 
      Appearance      =   0  'Flat
      Height          =   150
      Left            =   1035
      MouseIcon       =   "Cd Player1.frx":1978
      MousePointer    =   99  'Custom
      Picture         =   "Cd Player1.frx":1DBA
      Stretch         =   -1  'True
      ToolTipText     =   "时间显示方式  - Snowman Media  2.0"
      Top             =   1485
      Width           =   165
   End
   Begin VB.Image Image8 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      MouseIcon       =   "Cd Player1.frx":21B9
      MousePointer    =   99  'Custom
      Picture         =   "Cd Player1.frx":230B
      Stretch         =   -1  'True
      ToolTipText     =   "停止播放  - Snowman Media  2.0"
      Top             =   2340
      Width           =   300
   End
   Begin VB.Label TotalTrack 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   315
      TabIndex        =   3
      ToolTipText     =   "总时间  - Snowman Media  2.0"
      Top             =   3150
      Width           =   90
   End
   Begin VB.Label TrackTime 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   315
      TabIndex        =   2
      ToolTipText     =   "本首时间  - Snowman Media  2.0"
      Top             =   2925
      Width           =   90
   End
   Begin VB.Image Image14 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   0
      MouseIcon       =   "Cd Player1.frx":270A
      MousePointer    =   99  'Custom
      Picture         =   "Cd Player1.frx":2A14
      Stretch         =   -1  'True
      ToolTipText     =   "弹出功能菜单  - Snowman Media  2.0"
      Top             =   630
      Width           =   435
   End
   Begin VB.Image Image13 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1035
      MouseIcon       =   "Cd Player1.frx":2E13
      MousePointer    =   99  'Custom
      Picture         =   "Cd Player1.frx":2F65
      ToolTipText     =   "打开音频媒体  - Snowman Media  2.0"
      Top             =   1665
      Width           =   240
   End
   Begin VB.Image Image12 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1305
      MouseIcon       =   "Cd Player1.frx":3305
      MousePointer    =   99  'Custom
      Picture         =   "Cd Player1.frx":3457
      ToolTipText     =   "退出,谢谢使用  - Snowman Media  2.0"
      Top             =   1665
      Width           =   240
   End
   Begin VB.Image Image6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      MouseIcon       =   "Cd Player1.frx":381A
      MousePointer    =   99  'Custom
      Picture         =   "Cd Player1.frx":396C
      Stretch         =   -1  'True
      ToolTipText     =   "向后退  - Snowman Media  2.0"
      Top             =   2025
      Width           =   300
   End
   Begin VB.Image Image11 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1485
      MouseIcon       =   "Cd Player1.frx":3D6B
      MousePointer    =   99  'Custom
      Picture         =   "Cd Player1.frx":3EBD
      Stretch         =   -1  'True
      ToolTipText     =   "弹出光驱  - Snowman Media  2.0"
      Top             =   2340
      Width           =   300
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   360
      MouseIcon       =   "Cd Player1.frx":42BC
      MousePointer    =   99  'Custom
      Picture         =   "Cd Player1.frx":440E
      Stretch         =   -1  'True
      ToolTipText     =   "播放CD  - Snowman Media  2.0"
      Top             =   2025
      Width           =   255
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   720
      MouseIcon       =   "Cd Player1.frx":480D
      MousePointer    =   99  'Custom
      Picture         =   "Cd Player1.frx":495F
      Stretch         =   -1  'True
      ToolTipText     =   "暂停播放  - Snowman Media  2.0"
      Top             =   2025
      Width           =   255
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      MouseIcon       =   "Cd Player1.frx":4D5E
      MousePointer    =   99  'Custom
      Picture         =   "Cd Player1.frx":4EB0
      Stretch         =   -1  'True
      ToolTipText     =   "下一首  - Snowman Media  2.0"
      Top             =   2340
      Width           =   300
   End
   Begin VB.Image Image5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   315
      MouseIcon       =   "Cd Player1.frx":52AF
      MousePointer    =   99  'Custom
      Picture         =   "Cd Player1.frx":5401
      Stretch         =   -1  'True
      ToolTipText     =   "上一首  - Snowman Media  2.0"
      Top             =   2340
      Width           =   300
   End
   Begin VB.Image Image7 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1485
      MouseIcon       =   "Cd Player1.frx":5800
      MousePointer    =   99  'Custom
      Picture         =   "Cd Player1.frx":5952
      Stretch         =   -1  'True
      ToolTipText     =   "向前进  - Snowman Media  2.0"
      Top             =   2025
      Width           =   300
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2895
      Left            =   0
      Picture         =   "Cd Player1.frx":5D51
      ToolTipText     =   "嘘,静下来慢慢欣赏!  －- 流动网络H2ont祝你有愉悦的多媒体体验"
      Top             =   0
      Width           =   1845
   End
End
Attribute VB_Name = "Form102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SETHOTKEY = &H32
Private Const HOTKEYF_SHIFT = &H1
Private Const HOTKEYF_CONTROL = &H2
Private Const HOTKEYF_ALT = &H4
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim intpos As Integer
Const hwnd_top = 1
Dim ingreturnvalue As Long
Const HWND_TOPMOST = -1
Const SWP_SHOWWINDOW = &H40
Dim FastForwardSpeed As Long
Dim Playing As Boolean
Dim CDLoad As Boolean
Dim TotalTracks As Integer
Dim TrackLength() As String
Dim Track As Integer
Dim Minute As Integer
Dim Second As Integer
Dim Command As String
Dim hmixer As Long
Dim volCtrl As MIXERCONTROL
Private Function SendMCIString(Cmd As String, fShowError As Boolean) As Boolean
Static rc As Long
Static errStr As String * 400
rc = mciSendString(Cmd, 0, 0, hWnd)
'If (fShowError And rc <> 0) Then
 '   mciGetErrorString rc, errStr, Len(errStr)
  '  MsgBox errStr
'End If
SendMCIString = (rc = 0)
End Function
Private Sub Image12_Click()
SendMCIString "stop cd wait", True
Command = "seek cd to " & Track
SendMCIString Command, True
Playing = False
Update
SendMCIString "close all", False
Unload Me
End
End Sub
Private Sub Form_Load()
  Dim l As Long
   Dim wHotkey As Long
   wHotkey = (HOTKEYF_ALT Or HOTKEYF_CONTROL) * (2 ^ 8) + 65
   l = SendMessage(Me.hWnd, WM_SETHOTKEY, wHotkey, 0)
ingreturnvalue = SetWindowPos(Me.hWnd, HWND_TOPMOST, Val(10), Val(10), Val(10), Val(10), SWP_SHOWWINDOW)
    Form102.Width = 5235
    Form102.Height = 3105
Dim rc  As Long
Dim OK As Boolean
rc = mixerOpen(hmixer, 0, 0, 0, 0)
If MMSYSERR_NOERROR <> rc Then
    MsgBox "Could not open the mixer.", vbCritical, "Volume Control"
    Exit Sub
End If
OK = fGetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_VOLUME, volCtrl)
If OK Then
    With Volume
        .Max = volCtrl.lMinimum
        .Min = volCtrl.lMaximum \ 2
        .SmallChange = 1000
        .LargeChange = 1000
    End With
End If
    Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - Height) \ 2
Timer1.Interval = 500
If (App.PrevInstance = True) Then
    End
End If
Timer1.Enabled = False
FastForwardSpeed = 10
CDLoad = False
If (SendMCIString("open cdaudio alias cd wait shareable", True) = False) Then
    End
End If
SendMCIString "set cd time format tmsf wait", True
Timer1.Enabled = True
'MsgBox ("Open CD rom Drive.")
'SendMCIString "set cd door open", True
'MsgBox ("Put your compact disk in the CD Rom drive and click Close.")
Dim lRegion As Long
      If FileExist(App.Path & "\006.bmp") Then
        lRegion = TransparentForm(App.Path & "\006.bmp")
        Call SetWindowRgn(Me.hWnd, lRegion, True)
    End If
Form102.MediaPlayer1.filename = App.Path + "\start.wav"
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
outlook.Show
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SendMCIString "close all", False
End Sub
Private Sub Image14_Click()
outlook.Show
End Sub
Private Sub image2_Click()
SendMCIString "play cd", True
Playing = True
End Sub
Private Sub image3_Click()
SendMCIString "pause cd", True
Playing = False
Update
End Sub
Private Sub image11_Click()
If Form102.Eject.Enabled = True Then
SendMCIString "set cd door open", True
Else
 SendMCIString "set cd door closed", True
 End If
Update
End Sub
Private Sub image7_Click()
Dim e As String * 40
SendMCIString "set cd time format milliseconds", True
mciSendString "status cd position wait", e, Len(e), 0
If (Playing) Then
    Command = "play cd from " & CStr(CLng(e) + FastForwardSpeed * 1000)
Else
    Command = "seek cd to " & CStr(CLng(e) + FastForwardSpeed * 1000)
End If
mciSendString Command, 0, 0, 0
SendMCIString "set cd time format tmsf", True
Update
End Sub
Private Sub Image6_Click()
Dim e As String * 40
SendMCIString "set cd time format milliseconds", True
mciSendString "status cd position wait", e, Len(e), 0
If (Playing) Then
    Command = "play cd from " & CStr(CLng(e) - FastForwardSpeed * 1000)
Else
    Command = "seek cd to " & CStr(CLng(e) - FastForwardSpeed * 1000)
End If
mciSendString Command, 0, 0, 0
SendMCIString "set cd time format tmsf", True
Update
End Sub
Private Sub image4_Click()
If (Track < TotalTracks) Then
    If (Playing) Then
        Command = "play cd from " & Track + 1
        SendMCIString Command, True
    Else
        Command = "seek cd to " & Track + 1
        SendMCIString Command, True
    End If
Else
    SendMCIString "seek cd to 1", True
End If
Update
End Sub
Private Sub image5_Click()
Dim from As String
If (Minute = 0 And Second = 0) Then
    If (Track > 1) Then
        from = CStr(Track - 1)
    Else
        from = CStr(TotalTracks)
    End If
Else
    from = CStr(Track)
End If
If (Playing) Then
    Command = "play cd from " & from
    SendMCIString Command, True
Else
    Command = "seek cd to " & from
    SendMCIString Command, True
End If
Update
End Sub
Private Sub Update()
Static e As String * 30
mciSendString "status cd media present", e, Len(e), 0
If (CBool(e)) Then
    If (CDLoad = False) Then
        mciSendString "status cd number of tracks wait", e, Len(e), 0
        TotalTracks = CInt(Mid$(e, 1, 2))
        Eject.Enabled = True
        Form102.Image11.ToolTipText = "弹出光驱  - Snowman Media  2.0"
        If (TotalTracks = 1) Then
            Exit Sub
        End If
        mciSendString "status cd length wait", e, Len(e), 0
        TotalTrack.Caption = TotalTracks & "/" & e
        ReDim TrackLength(1 To TotalTracks)
        Dim i As Integer
        For i = 1 To TotalTracks
            Command = "status cd length track " & i
            mciSendString Command, e, Len(e), 0
            TrackLength(i) = e
        Next
        Dim ts As Integer
        TrackSelection.Clear
        For ts = 1 To TotalTracks
        TrackSelection.AddItem ts
        Next ts
        TrackSelection.Text = TrackSelection.List(0)
        CDLoad = True
        SendMCIString "seek cd to 1", True
    End If
     mciSendString "status cd position", e, Len(e), 0
    Track = CInt(Mid$(e, 1, 2))
    Minute = CInt(Mid$(e, 4, 2))
    Second = CInt(Mid$(e, 7, 2))
    TimeWindow.Text = "[" & Format(Track, "00") & "] " & Format(Minute, "00") _
            & ":" & Format(Second, "00")
             TrackTime.Caption = TrackLength(Track)
    TrackSelection.Text = TrackSelection.List(Track - 1)
      mciSendString "status cd mode", e, Len(e), 0
    Playing = (Mid$(e, 1, 7) = "playing")
Else
     Eject.Enabled = False
     Form102.Image11.ToolTipText = "送入光驱  - Snowman Media  2.0"
     If (CDLoad = True) Then
        CDLoad = False
        Playing = False
        TrackTime.Caption = ""
        TrackTime.Caption = ""
        TimeWindow.Text = ""
    End If
End If
End Sub
Private Sub image8_Click()
SendMCIString "stop cd wait", True
Command = "seek cd to " & Track
SendMCIString Command, True
Playing = False
Update
End Sub
Private Function fSetVolumeControl(ByVal hmixer As Long, _
    mxc As MIXERCONTROL, ByVal Volume As Long) As Boolean
Dim rc   As Long
Dim mxcd As MIXERCONTROLDETAILS
Dim vol  As MIXERCONTROLDETAILS_UNSIGNED
With mxcd
    .item = 0
    .dwControlID = mxc.dwControlID
    .cbStruct = Len(mxcd)
    .cbDetails = Len(vol)
End With
hmem = GlobalAlloc(&H40, Len(vol))
mxcd.paDetails = GlobalLock(hmem)
mxcd.cChannels = 1
vol.dwValue = Volume
Call CopyPtrFromStruct(mxcd.paDetails, vol, Len(vol))
rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
Call GlobalFree(hmem)
If MMSYSERR_NOERROR = rc Then
    fSetVolumeControl = True
Else
    fSetVolumeControl = False
End If
End Function
Private Sub Image9_Click()
a = a + 1
If a > 3 Then
a = 1
End If
If a = 1 Then
   TimeWindow.Top = 2700
   TrackTime.Top = 4700
   TotalTrack.Top = 4700
   Else
        If a = 2 Then
        TimeWindow.Top = 4700
        TrackTime.Top = 2700
        TotalTrack.Top = 4700
             Else
               If a = 3 Then
                TimeWindow.Top = 4700
                 TrackTime.Top = 4700
                   TotalTrack.Top = 2700
                End If
        End If
End If
End Sub

Private Sub TrackSelection_Click()
If (CDLoad) Then
        If (Track <= TotalTracks) Then
            If (Playing) Then
                Command = "play cd from " & Val(TrackSelection.Text)
                SendMCIString Command, True
             Else
                Command = "seek cd to " & Val(TrackSelection.Text)
                SendMCIString Command, True
                SendMCIString "play cd", True
                Playing = True
            End If
        End If
        Else
        SendMCIString "seek cd to 1", True
    End If
    Update
End Sub
Private Sub Volume_Change()
Dim lVol As Long
lVol = CLng(Volume.Value) * 2
Call fSetVolumeControl(hmixer, volCtrl, lVol)
End Sub
Private Sub Volume_Scroll()
Dim lVol As Long
lVol = CLng(Volume.Value) * 2
Call fSetVolumeControl(hmixer, volCtrl, lVol)
End Sub
Private Sub image1_mousedown(Button As Integer, shift As Integer, x As Single, y As Single)
If Button = 1 Then
a = x
b = y
End If
 If Button = 2 Then
 ingreturnvalue = SetWindowPos(Me.hWnd, hwnd_top, Val(10), Val(10), Val(10), Val(10), SWP_SHOWWINDOW)
     Form102.Width = 5235
     Form102.Height = 3105
     End If
 End Sub
Private Sub image1_mousemove(Button As Integer, shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub
c = x
d = y
Form102.Left = Form102.Left + (c - a)
Form102.Top = Form102.Top + (d - b)
End Sub
Private Sub Timer1_Timer()
  Update
End Sub
Private Sub image13_Click()
  Form102.CommonDialog1.Filter = "媒体文件:vcd、mp3、wma、wav、wax、asf、rmi、asx、mov、m1v、mp2、mpg、mpeg、mpa、mpe、avi、mid、qt、m3u、aif、aifc、aiff、au、snd..." & _
   "|*.au;*.dat;*.and;*.aif;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.m3u;*.wma;*.wav;*.avi;*.mid|浏览所有文件:*.*|*.*"
        Form102.CommonDialog1.FilterIndex = 1
    Form102.CommonDialog1.ShowOpen
    If Len(Form102.CommonDialog1.filename) > 0 Then
    Form6.MediaPlayer1.filename = Form102.CommonDialog1.filename
    Form6.Caption = Form102.CommonDialog1.filename & "  - Snowman Media  2.0"
    Form6.Show
End If
End Sub
