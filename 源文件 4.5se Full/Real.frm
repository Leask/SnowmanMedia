VERSION 5.00
Object = "{CFCDAA00-8BE4-11CF-B84B-0020AFBBCCFA}#1.0#0"; "rmoc3260.dll"
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "SmM_Tools.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "S.m.M. Real Media Player"
   ClientHeight    =   3690
   ClientLeft      =   3150
   ClientTop       =   3390
   ClientWidth     =   5880
   Icon            =   "Real.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5355
      Top             =   3825
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      FillColor       =   &H00C0FFFF&
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   90
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   2
      Top             =   4500
      Width           =   195
   End
   Begin API控制大全.LyfTools Ly 
      Left            =   4635
      Top             =   3825
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      FillColor       =   &H00C0FFFF&
      ForeColor       =   &H00C0FFFF&
      Height          =   105
      Left            =   45
      ScaleHeight     =   75
      ScaleWidth      =   5880
      TabIndex        =   3
      Top             =   4545
      Width           =   5910
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4515
      Begin VB.Image Image1 
         Height          =   1500
         Left            =   1350
         Top             =   315
         Width           =   2220
      End
   End
   Begin RealAudioObjectsCtl.RealAudio Rm 
      Height          =   4500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   7938
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
   Begin VB.Menu ddf 
      Caption         =   "控制(&C)"
      Begin VB.Menu dfg 
         Caption         =   "播放(&P)"
      End
      Begin VB.Menu dsf 
         Caption         =   "暂停(&A)"
      End
      Begin VB.Menu sdff 
         Caption         =   "停止(&S)"
      End
      Begin VB.Menu sdf 
         Caption         =   "-"
      End
      Begin VB.Menu erer 
         Caption         =   "倒退 5 秒(&B)"
      End
      Begin VB.Menu erzxc 
         Caption         =   "快进 5 秒(&F)"
      End
   End
   Begin VB.Menu hrth 
      Caption         =   "查看(&V)"
      Begin VB.Menu ry 
         Caption         =   "统计信息(&I)"
      End
   End
   Begin VB.Menu hery 
      Caption         =   "视图(&S)"
      Begin VB.Menu retv 
         Caption         =   "全屏(&F)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Text As String
Dim xx As Integer
Dim UN As Boolean








Private Sub dfg_Click()
Rm.DoPlay
End Sub

Private Sub dsf_Click()
Rm.DoPause
End Sub

Private Sub erer_Click()
If Rm.GetPosition > 5000 Then
Rm.SetPosition Rm.GetPosition - 5000
Else
Rm.SetPosition 0

End If
End Sub

Private Sub erzxc_Click()
If Rm.GetPosition > Rm.GetLength - 5000 Then Exit Sub

Rm.SetPosition Rm.GetPosition + 5000

End Sub

Private Sub Form_Load()

  Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "RealPlay", True
Rm.SetEnableContextMenu False
Rm.SetNoLogo True
Rm.SetMaintainAspect True
Rm.SetEnableMessageBox False
Image1.Picture = LoadPicture(Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "App_Path"))
End Sub

Private Sub Form_Resize()
On Error Resume Next

If Form1.WindowState = 1 Then Exit Sub
If Me.Width < 4000 Then Me.Width = 4000
If Me.Height < 3000 Then Me.Height = 3000
Rm.Visible = False
Rm.Width = Me.Width - 140
Rm.Height = Me.Height - 1030
Rm.Visible = True
Frame1.Width = Me.Width - 140
Frame1.Height = Me.Height - 1350
Image1.Left = (Image1.Width + Me.Width) / 2 - Image1.Width - 100
Image1.Top = (Image1.Height + Me.Height) / 2 - Image1.Height - 600


Picture2.Top = Me.Height - 970
Picture1.Top = Me.Height - 1020
Picture2.Width = Me.Width - 210
End Sub


Private Sub Form_Unload(Cancel As Integer)
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "RealPlay", False
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
xx = X
Rm.DoPlay

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button <> 1 Then Exit Sub
  If Picture1.Left + X - xx < 90 Then Picture1.Left = 90
  If Picture1.Left + X - xx > Picture2.Width - 200 Then Picture1.Left = Picture2.Width - 200
 If Picture1.Left + X - xx >= 90 And Picture1.Left + X - xx <= Picture2.Width - 200 Then Picture1.Left = Picture1.Left + X - xx

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
 If Button <> 1 Then Exit Sub
 
 
 
 
 Rm.SetPosition (Picture1.Left - 100) * Rm.GetLength / (Picture2.Width - 300)

End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
 If Button <> 1 Then Exit Sub
 Picture1.Left = X
 Rm.SetPosition (Picture1.Left - 100) * Rm.GetLength / (Picture2.Width - 300)
Rm.DoPlay

End Sub

Private Sub retv_Click()
Rm.SetFullScreen
End Sub

Public Function Gtime(Value As Long) As String
On Error Resume Next
Dim SS, FF, mm, FM As Long
Dim SSD, FFD, mmD As String
If Int(Value) < 0 Then Exit Function
SS = Int(Int(Value) / 3600)
FF = Int(Int(Value) / 60 - SS * 60)
FM = Int(Int(Value) / 60)
mm = Int(Value) - 60 * FM
SSD = SS
If SS < 10 Then SSD = "0" & SS
FFD = FF
If FF < 10 Then FFD = "0" & FF
mmD = mm
If mm < 10 Then mmD = "0" & mm
If SS = 0 Then
Gtime = FFD & ":" & mmD
Else
Gtime = SSD & ":" & FFD & ":" & mmD
End If
End Function


Private Sub Rm_OnPlayStateChange(ByVal lNewState As Long)
If lNewState = 0 Then UN = True
If Rm.GetClipWidth > 0 Then
Frame1.Visible = False
Else
Frame1.Visible = True

End If
End Sub

Private Sub Rm_OnPositionChange(ByVal lPos As Long, ByVal lLen As Long)
On Error Resume Next
Picture1.Left = Rm.GetPosition / Rm.GetLength * (Picture2.Width - 300) + 100
End Sub


Private Sub ry_Click()
Rm.SetShowStatistics True
End Sub

Private Sub sdff_Click()
Unload Me
End
End Sub

Private Sub Timer1_Timer()

On Error Resume Next
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "RealPlay") = True Then

 Rm.Source = myReadINI(App.Path + "\SmM_RealMedia.dat", "Real", "Filename", "")
   Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "RealPlay", False
 If Rm.GetClipWidth > 0 Then
   Rm.Visible = False
   Rm.Width = Rm.GetClipWidth * Screen.TwipsPerPixelX
   Me.Width = Rm.Width + 140
   Rm.Height = Rm.GetClipHeight * Screen.TwipsPerPixelY
   Rm.Visible = True
   Me.Height = Rm.Height + 1030
  Else
   Me.Width = 6000
   Me.Height = 4500
   Form_Resize
  End If
 
End If

If Len(Rm.Source) > 0 Then

Text = Gtime(Rm.GetPosition / 1000) + " / " + Left(Gtime(Rm.GetLength / 1000), 5) + " - " + Right(Str(Int(Rm.GetBandwidthCurrent / 1000)), Len(Str(Int(Rm.GetBandwidthCurrent / 1000))) - 1) + "Kbps  "
If Len(Rm.GetTitle) > 0 Then Text = Text + "标题:" + Rm.GetTitle + "  "
If Len(Rm.GetAuthor) > 0 Then Text = Text + "艺术家:" + Rm.GetAuthor + "  "
If Len(Rm.GetCopyright) > 0 Then Text = Text + "  版权:" + Rm.GetCopyright + "  "

If Rm.GetClipWidth > 0 Then
Text = Text + "视频:" + Str(Rm.GetClipWidth) + " ×" + Str(Rm.GetClipHeight) + "  "
Else
Text = Text + "类型:音频  "
End If

If Left(Rm.Source, 7) = "file://" Then
Me.Caption = Text + "地址:" + Right(Rm.Source, Len(Rm.Source) - 7)
Else
Me.Caption = Text + "地址:" + Rm.Source
End If

Else
Me.Caption = "S.m.M. Real Media Player"
End If

If App.PrevInstance = True Or UN = True Then Unload Me

End Sub
