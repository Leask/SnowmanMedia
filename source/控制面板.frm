VERSION 5.00
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sm.M. Net Show"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "控制面板.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   8055
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   135
      Picture         =   "控制面板.frx":1582
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   34
      Top             =   90
      Width           =   510
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   -45
      TabIndex        =   32
      Top             =   675
      Width           =   735
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         FillColor       =   &H80000008&
         Height          =   1725
         Left            =   -90
         Top             =   0
         Width           =   825
      End
   End
   Begin ACTIVESKINLibCtl.SkinForm SkinForm1 
      Height          =   480
      Left            =   3420
      OleObjectBlob   =   "控制面板.frx":188C
      TabIndex        =   31
      Top             =   1800
      Width           =   480
   End
   Begin VB.PictureBox Picture15 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6345
      ScaleHeight     =   255
      ScaleWidth      =   1515
      TabIndex        =   29
      Top             =   3240
      Width           =   1545
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "隐藏高级选项(&H)"
         ForeColor       =   &H00FF0000&
         Height          =   1185
         Left            =   45
         TabIndex        =   30
         Top             =   45
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   135
      ScaleHeight     =   885
      ScaleWidth      =   3810
      TabIndex        =   26
      Top             =   2610
      Width           =   3840
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "注意:"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   90
         TabIndex        =   28
         Top             =   45
         Width           =   450
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "以下选项同等于Windows系统上的相关设置,它们可能与其它有关程序共享.当以下设置改动时其它设置将随即被改动。"
         ForeColor       =   &H00FF0000&
         Height          =   915
         Left            =   90
         TabIndex        =   27
         Top             =   270
         Width           =   3705
      End
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   810
      ScaleHeight     =   840
      ScaleWidth      =   7050
      TabIndex        =   23
      Top             =   90
      Width           =   7080
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "你还没有连接上Internet,进行网络在线播放之前必须先连接上Internet."
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   135
         TabIndex        =   25
         Top             =   90
         Width           =   6585
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   $"控制面板.frx":18D5
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   135
         TabIndex        =   24
         Top             =   360
         Width           =   6810
      End
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   810
      ScaleHeight     =   255
      ScaleWidth      =   2460
      TabIndex        =   21
      Top             =   1440
      Width           =   2490
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "显示高级选项(&S)"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   45
         TabIndex        =   22
         Top             =   45
         Width           =   2400
      End
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6795
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   10
      Top             =   1935
      Width           =   1095
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "取消(&C)"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   45
         TabIndex        =   20
         Top             =   45
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5580
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   9
      Top             =   1935
      Width           =   1095
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "确定(&Y)"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   45
         TabIndex        =   19
         Top             =   45
         Width           =   1680
      End
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4095
      ScaleHeight     =   255
      ScaleWidth      =   2190
      TabIndex        =   8
      Top             =   3240
      Width           =   2220
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "网络设置(&N)"
         ForeColor       =   &H00FF0000&
         Height          =   780
         Left            =   45
         TabIndex        =   18
         Top             =   45
         Width           =   2130
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6345
      ScaleHeight     =   255
      ScaleWidth      =   1515
      TabIndex        =   7
      Top             =   2925
      Width           =   1545
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "调制解调器(&M)"
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   45
         TabIndex        =   17
         Top             =   45
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4095
      ScaleHeight     =   255
      ScaleWidth      =   2190
      TabIndex        =   6
      Top             =   2925
      Width           =   2220
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "声音及多媒体(&S)"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   45
         TabIndex        =   16
         Top             =   45
         Width           =   1725
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6345
      ScaleHeight     =   255
      ScaleWidth      =   1515
      TabIndex        =   5
      Top             =   2610
      Width           =   1545
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "Internet选项(&I)"
         ForeColor       =   &H00FF0000&
         Height          =   510
         Left            =   45
         TabIndex        =   15
         Top             =   45
         Width           =   3210
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3420
      ScaleHeight     =   255
      ScaleWidth      =   4440
      TabIndex        =   4
      Top             =   1440
      Width           =   4470
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "不拨号连接在局域网或本地主机上使用网络播放(&T)"
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   45
         TabIndex        =   14
         Top             =   45
         Width           =   4830
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3420
      ScaleHeight     =   255
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   1080
      Width           =   2265
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "指定连接连接Internet(&P)"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   45
         TabIndex        =   13
         Top             =   45
         Width           =   2895
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4095
      ScaleHeight     =   255
      ScaleWidth      =   2190
      TabIndex        =   2
      Top             =   2610
      Width           =   2220
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "新建连接连接Internet(&N)"
         ForeColor       =   &H00FF0000&
         Height          =   465
         Left            =   45
         TabIndex        =   12
         Top             =   45
         Width           =   3705
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   810
      ScaleHeight     =   255
      ScaleWidth      =   2460
      TabIndex        =   1
      Top             =   1080
      Width           =   2490
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "默认连接连接Internet(&A)"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   45
         TabIndex        =   11
         Top             =   45
         Width           =   2985
      End
   End
   Begin VB.TextBox Connection 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5805
      TabIndex        =   0
      Top             =   1080
      Width           =   2085
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   3660
      Left            =   5355
      TabIndex        =   33
      Top             =   1845
      Width           =   4110
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   4200
         Left            =   0
         Top             =   0
         Width           =   5010
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   5685
      Left            =   -90
      TabIndex        =   35
      Top             =   -225
      Width           =   9915
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
 SkinForm1.SkinPath = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path")
 If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Op_" + Str(8)) = True Then
Call Label8_Click
Exit Sub
 End If
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Op_" + Str(9)) = True Then
Call Label5_Click
 Exit Sub
End If
 If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Op_" + Str(10)) = True Then
    Connection.Text = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Te_" + Str(1))
Call Label7_Click
End If
 End Sub



Private Sub Form_Unload(Cancel As Integer)
       Set Form2 = Nothing

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture1.BackColor = &HFFFF&
Picture2.BackColor = &HFFFF&
Picture4.BackColor = &HFFFF&
Picture3.BackColor = &HFFFF&
Picture5.BackColor = &HFFFF&
Picture6.BackColor = &HFFFF&
Picture7.BackColor = &HFFFF&
Picture8.BackColor = &HFFFF&
Picture9.BackColor = &HFFFF&
Picture10.BackColor = &HFFFF&
Picture11.BackColor = &HFFFF&
Picture15.BackColor = &HFFFF&
Label5.BackColor = &HFFFF&
Label5.ForeColor = &HFF0000
Label6.BackColor = &HFFFF&
Label6.ForeColor = &HFF0000
Label7.BackColor = &HFFFF&
Label7.ForeColor = &HFF0000
Label8.BackColor = &HFFFF&
Label8.ForeColor = &HFF0000
Label9.BackColor = &HFFFF&
Label9.ForeColor = &HFF0000
Label10.BackColor = &HFFFF&
Label10.ForeColor = &HFF0000
Label11.BackColor = &HFFFF&
Label11.ForeColor = &HFF0000
Label12.BackColor = &HFFFF&
Label12.ForeColor = &HFF0000
Label13.BackColor = &HFFFF&
Label13.ForeColor = &HFF0000
Label14.BackColor = &HFFFF&
Label14.ForeColor = &HFF0000
Label15.BackColor = &HFFFF&
Label15.ForeColor = &HFF0000
Label16.BackColor = &HFFFF&
Label16.ForeColor = &HFF0000
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture1.BackColor = &HFFFF&
Picture2.BackColor = &HFFFF&
Picture4.BackColor = &HFFFF&
Picture3.BackColor = &HFFFF&
Picture5.BackColor = &HFFFF&
Picture6.BackColor = &HFFFF&
Picture7.BackColor = &HFFFF&
Picture8.BackColor = &HFFFF&
Picture9.BackColor = &HFFFF&
Picture10.BackColor = &HFFFF&
Picture11.BackColor = &HFFFF&
Picture15.BackColor = &HFFFF&
Label5.BackColor = &HFFFF&
Label5.ForeColor = &HFF0000
Label6.BackColor = &HFFFF&
Label6.ForeColor = &HFF0000
Label7.BackColor = &HFFFF&
Label7.ForeColor = &HFF0000
Label8.BackColor = &HFFFF&
Label8.ForeColor = &HFF0000
Label9.BackColor = &HFFFF&
Label9.ForeColor = &HFF0000
Label10.BackColor = &HFFFF&
Label10.ForeColor = &HFF0000
Label11.BackColor = &HFFFF&
Label11.ForeColor = &HFF0000
Label12.BackColor = &HFFFF&
Label12.ForeColor = &HFF0000
Label13.BackColor = &HFFFF&
Label13.ForeColor = &HFF0000
Label14.BackColor = &HFFFF&
Label14.ForeColor = &HFF0000
Label15.BackColor = &HFFFF&
Label15.ForeColor = &HFF0000
Label16.BackColor = &HFFFF&
Label16.ForeColor = &HFF0000
End Sub

Private Sub Label10_Click()
ControlPanels ("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,,0")
End Sub
Private Sub Label11_Click()
ControlPanels ("rundll32.exe shell32.dll,Control_RunDLL modem.cpl")
End Sub
Private Sub Label12_Click()
ControlPanels ("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl")
End Sub
Private Sub Label13_Click()
Unload Me
End Sub
Private Sub Label14_Click()
Unload Me
End Sub
Private Sub Label15_Click()
Me.Height = 4170
End Sub
Private Sub Label16_Click()
Me.Height = 2850
End Sub
Private Sub Label5_Click()
 Dim rtn
  rtn = Shell("rundll32.exe rnaui.dll,RnaDial " & _
          0)
          Formo.Show
          Unload Me
End Sub
Private Sub Label6_Click()
 Dim rtn
 rtn = Shell("rundll32.exe rnaui.dll,RnaWizard ", 0)
End Sub
Private Sub Label7_Click()
Dim rtn
    rtn = Shell("rundll32.exe rnaui.dll,RnaDial " & _
          Connection.Text, 0)
          Formo.Show
          Unload Me
End Sub

Private Sub Label8_Click()
Formo.Show
Unload Me
End Sub
Private Sub Label9_Click()
ControlPanels ("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0")
End Sub
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label7.BackColor = &HFF0000
Picture3.BackColor = &HFF0000
Label7.ForeColor = &HFFFF&
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5.BackColor = &HFF0000
Picture1.BackColor = &HFF0000
Label5.ForeColor = &HFFFF&
End Sub
Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label15.BackColor = &HFF0000
Picture11.BackColor = &HFF0000
Label15.ForeColor = &HFFFF&
End Sub
Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label8.BackColor = &HFF0000
Picture4.BackColor = &HFF0000
Label8.ForeColor = &HFFFF&
End Sub
Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label13.BackColor = &HFF0000
Picture9.BackColor = &HFF0000
Label13.ForeColor = &HFFFF&
End Sub
Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label6.BackColor = &HFF0000
Picture2.BackColor = &HFF0000
Label6.ForeColor = &HFFFF&
End Sub
Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label9.BackColor = &HFF0000
Picture5.BackColor = &HFF0000
Label9.ForeColor = &HFFFF&
End Sub
Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label10.BackColor = &HFF0000
Picture6.BackColor = &HFF0000
Label10.ForeColor = &HFFFF&
End Sub
Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label11.BackColor = &HFF0000
Picture7.BackColor = &HFF0000
Label11.ForeColor = &HFFFF&
End Sub
Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label12.BackColor = &HFF0000
Picture8.BackColor = &HFF0000
Label12.ForeColor = &HFFFF&
End Sub
Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label16.BackColor = &HFF0000
Picture15.BackColor = &HFF0000
Label16.ForeColor = &HFFFF&
End Sub
Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label14.BackColor = &HFF0000
Picture10.BackColor = &HFF0000
Label14.ForeColor = &HFFFF&
End Sub
