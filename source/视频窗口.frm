VERSION 5.00
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Object = "{CFCDAA00-8BE4-11CF-B84B-0020AFBBCCFA}#1.0#0"; "RMOC3260.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form200 
   Caption         =   "Sm.M. Video Window"
   ClientHeight    =   4515
   ClientLeft      =   6150
   ClientTop       =   165
   ClientWidth     =   6765
   Icon            =   "视频窗口.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4515
   ScaleWidth      =   6765
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3555
      Top             =   6165
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2070
      Top             =   5985
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   5460
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7665
      Begin MSComctlLib.Slider Sld 
         Height          =   315
         Left            =   10000
         TabIndex        =   4
         Top             =   4320
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   556
         _Version        =   393216
         SelectRange     =   -1  'True
         TickStyle       =   3
      End
      Begin RealAudioObjectsCtl.RealAudio RA 
         Height          =   4380
         Left            =   10000
         TabIndex        =   3
         Top             =   -45
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   7726
         AUTOSTART       =   -1  'True
         SHUFFLE         =   0   'False
         PREFETCH        =   0   'False
         NOLABELS        =   0   'False
         CONTROLS        =   "imagewindow"
         LOOP            =   0   'False
         NUMLOOP         =   0
         CENTER          =   0   'False
         MAINTAINASPECT  =   0   'False
         BACKGROUNDCOLOR =   "#000000"
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   2265
         Left            =   1845
         Top             =   945
         Width           =   2805
      End
      Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
         Height          =   4965
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   6810
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
         ClickToPlay     =   0   'False
         CursorType      =   0
         CurrentPosition =   -1
         CurrentMarker   =   0
         DefaultFrame    =   ""
         DisplayBackColor=   0
         DisplayForeColor=   16777215
         DisplayMode     =   0
         DisplaySize     =   4
         Enabled         =   -1  'True
         EnableContextMenu=   0   'False
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
         SendMouseClickEvents=   -1  'True
         SendMouseMoveEvents=   -1  'True
         SendPlayStateChangeEvents=   -1  'True
         ShowCaptioning  =   0   'False
         ShowControls    =   -1  'True
         ShowAudioControls=   0   'False
         ShowDisplay     =   0   'False
         ShowGotoBar     =   0   'False
         ShowPositionControls=   0   'False
         ShowStatusBar   =   0   'False
         ShowTracker     =   -1  'True
         TransparentAtStart=   0   'False
         VideoBorderWidth=   0
         VideoBorderColor=   0
         VideoBorder3D   =   0   'False
         Volume          =   -600
         WindowlessVideo =   0   'False
      End
   End
   Begin ACTIVESKINLibCtl.SkinForm SkinForm1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "视频窗口.frx":1582
      TabIndex        =   2
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "Form200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RmGn As Boolean
Dim RMStop As Boolean








Private Sub Form_Resize()
On Error Resume Next

MediaPlayer1.Height = Me.Height - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "h", "") - 6345)
MediaPlayer1.Width = Me.Width - 100 - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "w", "") - 6240)
Frame1.Height = Me.Height
Frame1.Width = Me.Width
Image1.Top = Me.Height / 2 - Me.Image1.Height / 2 - 350
Image1.Left = Me.Width / 2 - Me.Image1.Width / 2 - 80
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(6)) = 1 Then
If Me.WindowState = vbMaximized Then
Me.WindowState = 1
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(10)) = 1 Then Form102.LyfTools1.MakeTop Form200, False
MediaPlayer1.DisplaySize = mpFullScreen
RA.SetFullScreen
Timer1.Enabled = True
End If
End If
If Len(RA.Source) > 0 Then  'And RA.GetPosition > 0 Then
RA.Visible = False
If RA.GetClipWidth / RA.GetClipHeight >= (Me.Width - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "w", "") - 6240) - 120) / (Me.Height - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "h", "") - 6345) - 670) Then
RA.Width = Me.Width - 120 - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "w", "") - 6240)
RA.Height = (Me.Width - 120 - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "w", "") - 6240)) * RA.GetClipHeight / RA.GetClipWidth
RA.Top = (Me.Height - 670 - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "h", "") - 6345) + RA.Height) / 2 - RA.Height - 50
RA.Left = (Me.Width - 120 - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "w", "") - 6240) + RA.Width) / 2 - RA.Width
Else
RA.Height = Me.Height - 670 - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "h", "") - 6345)
RA.Width = (Me.Height - 670 - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "h", "") - 6345)) * RA.GetClipWidth / RA.GetClipHeight
RA.Top = (Me.Height - 670 - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "h", "") - 6345) + RA.Height) / 2 - RA.Height - 50
RA.Left = (Me.Width - 120 - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "w", "") - 6240) + RA.Width) / 2 - RA.Width
End If
RA.Visible = True
Frame1.Height = Me.Height
Frame1.Width = Me.Width
Sld.Left = -45
Sld.Top = Me.Height - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "h", "") - 6345) - 650
Sld.Width = Me.Width - 30 - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "w", "") - 6240)
End If
If Len(MediaPlayer1.Filename) > 0 Then
Sld.Left = 10000
RA.Left = 10000
End If

End Sub
Private Sub Form_Load()
On Error Resume Next
   If Form102.FileExists(Form102.Label3.Caption + "\SmM_PT\SmM_LG.gif") = False Then End
          Image1.Picture = LoadPicture(Form102.Label3.Caption + "\SmM_PT\smM_LG.gif")
RA.SetNoLogo True
  RA.SetEnableContextMenu False
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(12)) = 1 Then MediaPlayer1.ClickToPlay = True
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(10)) = 1 Then
Form102.LyfTools1.MakeTop Form200, True
Form102.OnTop.Checked = True
Else
Form102.OnTop.Checked = False
End If

'SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
SkinForm1.ScanControls = 1
 SkinForm1.SkinPath = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path")
Timer1.Enabled = False
RmGn = True
RMStop = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(8)) = 1 Then
Form102.LyfTools1.MakeTop Form102, True
Form102.OnTop.Checked = True
Else
Form102.LyfTools1.MakeTop Form102, False
Form102.OnTop.Checked = False

End If
If Len(MediaPlayer1.Filename) > 0 Then
        
        Form102.jn = MediaPlayer1.Filename
        Form102.jd = MediaPlayer1.CurrentPosition
        MediaPlayer1.Filename = "LLXX"
        Form102.MediaPlayer1.Filename = Form102.jn
       Form102.MediaPlayer1.CurrentPosition = Form102.jd
      End If
If Len(RA.Source) > 0 And RA.GetPosition > 0 Then
        Form102.jn = RA.Source
        Form102.jd = RA.GetPosition
        RMStop = False
        RA.DoStop
         Form102.RMStop = False
         Form102.RA.DoStop
        Form102.RA.Source = Form102.jn
         Form102.RMStop = False
        Form102.RA.DoPlay
       Form102.RA.SetPosition Form102.jd


           
End If

Form102.WindowState = 0
Form102.Show

Form102.f200 = 0
  Set Form200 = Nothing
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu Form102.MumA, 0, x + Frame1.Left + Image1.Left, y + Frame1.Top + Image1.Top
If Button = 1 Then
Form102.MoveX = x
Form102.MoveY = y
End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub
Form200.Left = Form200.Left + (x - Form102.MoveX)
Form200.Top = Form200.Top + (y - Form102.MoveY)
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
 On Error Resume Next
 MediaPlayer1.Filename = "LLXX"
 Call Form102.mpn
 Form102.jn = Form102.MediaPlayer1.Filename
  Form102.MediaPlayer1.Filename = "LLXX"
 MediaPlayer1.Filename = Form102.jn
If Len(MediaPlayer1.Filename) = 0 Then
If Len(Form102.RA.Source) > 0 Then 'And Form102.RA.GetPosition > 0 Then
         Form102.jn = Form102.RA.Source
 'If Left(Form102.jn, 3) = "htt" Then Form102.jn = StCr(Form102.RA.Source, "http://", "")
 'If Left(Form102.jn, 3) = "ftp" Then Form102.jn = StCr(Form102.RA.Source, "ftp://", "")
 'If Left(Form102.jn, 3) = "fil" Then Form102.jn = StCr(Form102.RA.Source, "file://", "")
     
      Form102.RMStop = False
      Form102.RA.DoStop
      RA.Source = Form102.jn
      RA.DoPlay
Else
 Unload Me
End If
End If
  Call Form_Resize
End Sub

Private Sub MediaPlayer1_Error()
On Error Resume Next
If RA.Left = 0 Then Form102.RMStop = False
If RA.Left = 0 Then RMStop = False
Dim T As String
'If MediaPlayer1.ErrorCode = -2147220891 Then
    T = UCase(Right(Form102.ListFile.List(Form102.pid), 3))
     If T = ".RA" Or T = ".RM" Or T = "RAM" Or T = ".RT" Or T = ".RP" Or T = "SMI" Or T = "MIL" Then
        If RA.Source = "file://" + Form102.ListFile.List(Form102.pid) Or RA.Source = "http://" + Form102.ListFile.List(Form102.pid) Or RA.Source = "ftp://" + Form102.ListFile.List(Form102.pid) Or RA.Source = Form102.ListFile.List(Form102.pid) Then
            RA.SetPosition 0
            Exit Sub
        Else
           RA.Source = Form102.ListFile.List(Form102.pid)
        End If
     End If
'End If
  Call Form_Resize
End Sub

Private Sub MediaPlayer1_MouseDown(Button As Integer, ShiftState As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu Form102.MumA, 0, x + Frame1.Left, y + Frame1.Top
If Button = 1 Then
Form102.MoveX = x
Form102.MoveY = y
End If

End Sub

Private Sub MediaPlayer1_MouseMove(Button As Integer, ShiftState As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub
Form200.Left = Form200.Left + (x - Form102.MoveX)
Form200.Top = Form200.Top + (y - Form102.MoveY)

End Sub

Private Sub MediaPlayer1_NewStream()
Call Form_Resize
End Sub

Private Sub Timer1_Timer()
If MediaPlayer1.DisplaySize <> mpFullScreen And RA.GetFullScreen <> True Then
RA.SetFullScreen
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(10)) = 1 Then Form102.LyfTools1.MakeTop Form200, True
Timer1.Enabled = False
Me.WindowState = 0
Form200.Show

End If
End Sub

Private Sub Timer2_Timer()
If Form102.f200 = 1 Then Form102.Hide
If MediaPlayer1.ImageSourceWidth = 0 And RA.GetClipWidth = 0 Then Unload Me
End Sub
Private Sub RA_OnPositionChange(ByVal lPos As Long, ByVal lLen As Long)
If RmGn = True Then Sld.Value = RA.GetPosition
End Sub
Private Sub RA_OnClipClosed()
'Form102.RMStop = True
'Call Form102.RA_OnClipClosed'Unload Me
On Error Resume Next
RA.Left = 10000
Sld.Left = 10000
 If RMStop = True Then
 If Form102.pid + 1 = Form102.ListFile.ListCount And Form102.Check1.Value = 0 And Form102.Check2.Value = 0 And Form102.Label26.Caption <> "locked" Then
 Unload Me
 Exit Sub
 End If
 MediaPlayer1.Filename = "LLXX"
 Call Form102.mpn
 Form102.jn = Form102.MediaPlayer1.Filename
  Form102.MediaPlayer1.Filename = "LLXX"
 MediaPlayer1.Filename = Form102.jn
If Len(MediaPlayer1.Filename) = 0 Then
If Len(Form102.RA.Source) > 0 Then 'And Form102.RA.GetPosition > 0 Then
         Form102.jn = Form102.RA.Source
' If Left(Form102.jn, 3) = "htt" Then Form102.jn = StCr(Form102.RA.Source, "http://", "")
' If Left(Form102.jn, 3) = "ftp" Then Form102.jn = StCr(Form102.RA.Source, "ftp://", "")
' If Left(Form102.jn, 3) = "fil" Then Form102.jn = StCr(Form102.RA.Source, "file://", "")
     
      Form102.RMStop = False
      Form102.RA.DoStop
      RA.Source = Form102.jn
      RA.DoPlay
Else
 Unload Me
End If
End If
End If
RMStop = True
End Sub


Private Sub RA_OnClipOpened(ByVal shortClipName As String, ByVal url As String)
'Call Form_Resize
Sld.Max = RA.GetLength
Form102.Label1.Caption = "rm"
Call Form_Resize
End Sub
Private Sub SLD1_Click()
RmGn = True
End Sub

Private Sub SLD_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RmGn = False
End Sub

Private Sub SLD_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RA.SetPosition Sld.Value
RmGn = True
End Sub
Function StCr(SearchLine As String, SearchFor As String, ReplaceWith As String)
Dim vSearchLine As String, found As Integer

found = InStr(SearchLine, SearchFor): vSearchLine = SearchLine
If found <> 0 Then
vSearchLine = ""
If found > 1 Then vSearchLine = Left(SearchLine, found - 1)
vSearchLine = vSearchLine + ReplaceWith
If found + Len(SearchFor) - 1 < Len(SearchLine) Then _
vSearchLine = vSearchLine + Right$(SearchLine, Len(SearchLine) - found - Len(SearchFor) + 1)
End If
StCr = vSearchLine

End Function

