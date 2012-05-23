VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Begin VB.Form Formo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sm.M. Net Show"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "网络播放.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6270
   StartUpPosition =   2  '屏幕中心
   Begin ACTIVESKINLibCtl.SkinForm SkinForm1 
      Height          =   480
      Left            =   945
      OleObjectBlob   =   "网络播放.frx":1582
      TabIndex        =   15
      Top             =   2790
      Width           =   480
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   810
      ScaleHeight     =   255
      ScaleWidth      =   2685
      TabIndex        =   13
      Top             =   1530
      Width           =   2715
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "你想要打开的网络媒体地址:"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   45
         TabIndex        =   14
         Top             =   45
         Width           =   2610
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   135
      Picture         =   "网络播放.frx":15CB
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   90
      Width           =   510
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      FillColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   810
      ScaleHeight     =   1335
      ScaleWidth      =   5295
      TabIndex        =   8
      Top             =   90
      Width           =   5325
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "在线播放过程中无需保存下载的文件即可在当前状态下打开,前提是你必须已经连接到Internet上."
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   135
         TabIndex        =   11
         Top             =   900
         Width           =   4995
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "当你输入媒体文件的地址并按下""确定""后 Snowman Media ilxz 3.5 将为你打开它."
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   135
         TabIndex        =   10
         Top             =   45
         Width           =   5010
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "例如以下URL: http://www.h2ont.com/xx.wmv 或: http://www.h2ont.com/swf"
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   135
         TabIndex        =   9
         Top             =   450
         Width           =   4995
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "浏览(&B)..."
         ForeColor       =   &H00FF0000&
         Height          =   600
         Left            =   45
         TabIndex        =   7
         Top             =   45
         Width           =   1815
      End
   End
   Begin VB.Frame fraSep 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   45
      Left            =   810
      TabIndex        =   5
      Top             =   2385
      Width           =   9495
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2700
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "确定(&Y)"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   45
         TabIndex        =   4
         Top             =   45
         Width           =   1680
      End
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3870
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "取消(&C)"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   45
         TabIndex        =   2
         Top             =   45
         Width           =   1635
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7155
      Top             =   450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Net Show Opening Window  - Snowman Media  3.0"
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   810
      TabIndex        =   0
      ToolTipText     =   "在这里输入你想打开的网络媒体地址  － Snowman Media  3.0"
      Top             =   1935
      Width           =   5325
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   135
      TabIndex        =   16
      Top             =   675
      Width           =   510
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         Height          =   2310
         Left            =   0
         Top             =   0
         Width           =   510
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   4920
      Left            =   -360
      TabIndex        =   17
      Top             =   -360
      Width           =   7260
   End
End
Attribute VB_Name = "Formo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
       Dim idc As Integer
Private Sub Form_Load()
 SkinForm1.SkinPath = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path")

End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set Formo = Nothing
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture1.BackColor = &HFFFF&
Picture10.BackColor = &HFFFF&
Picture9.BackColor = &HFFFF&
Label13.BackColor = &HFFFF&
Label13.ForeColor = &HFF0000
Label14.BackColor = &HFFFF&
Label14.ForeColor = &HFF0000
Label5.BackColor = &HFFFF&
Label5.ForeColor = &HFF0000
End Sub

Private Sub Label13_Click()
Dim Filename As String
Filename = Text1.Text
If Len(Filename) > 0 Then
  If Form102.Label2.Caption = "media" Then
            For idc = 0 To Form102.ListFile.ListCount - 1
            If Form102.ListFile.List(idc) = Filename Then
            Form102.pid = idc
            Form102.ListFile.ListIndex = Form102.pid
            Exit For
            End If
            Next
            If Form102.pid = -1 Then
          Form102.ListFile.AddItem (Filename)
          Form102.pid = Form102.ListFile.ListCount - 1
          Form102.ListFile.ListIndex = Form102.pid
          End If
          Form102.MediaPlayer1.Filename = Form102.ListFile.List(Form102.pid)
           Form102.Label1.Caption = "media"
           Form102.TrackSelection.Left = 10000
           Form102.Frame8.Left = 2880
           Form102.Frame1.Left = 10000
        Unload Me
 End If
If Form102.Label2.Caption = "flash" Then
                    If Form102.FileExists(Label3.Caption + "\SmM_FP.exe") = True Then
         Shell Form102.Label3.Caption + "\SmM_FP.exe " + Filename, vbNormalFocus

       Else
         MsgBox ("找不到文件[ " + Label3.Caption + "\SmM_FP.exe" + " ].该文件可能已经丢失或被移动,请重新安装 Snowman Media ilxz 3.5")
       End If

                       Unload Me
        End If
End If
End Sub
Private Sub Label14_Click()
Unload Me
End Sub
Private Sub Label5_Click()
If Form102.Label2.Caption = "media" Then
          Form102.pid = -1
   
            CommonDialog1.Filter = "媒体文件:Mp3、Wma、Wmv、Wav、Wax、Ra、Rm、Asf、Rmi、Asx、Mov、M1v、Mp2、Mpg、Mpeg、Mpa、Mpe、Avi、Mid、Qt、Aif、Aifc、Aiff、Au、Snd、Smi、Smil、Rt、Mpv、Rp、Ram、Rmm、Rtx" & _
          "|*.au;*.and;*.aif;*.wmv;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.wma;*.wav;*.avi;*.mid;*.smi;*.smil;*.rt;*.mpv;*.rp;*.ram;*.rmm;*.rtx;*.ra;*.rm|所有文件:*.*|*.*"
          CommonDialog1.FilterIndex = 1
          CommonDialog1.Filename = ""
          CommonDialog1.ShowOpen
             If Len(CommonDialog1.Filename) > 0 Then
            For idc = 0 To Form102.ListFile.ListCount - 1
            If Form102.ListFile.List(idc) = CommonDialog1.Filename Then
            Form102.pid = idc
            Form102.ListFile.ListIndex = Form102.pid
            Exit For
            End If
            Next
            If Form102.pid = -1 Then
          Form102.ListFile.AddItem (CommonDialog1.Filename)
          Form102.pid = Form102.ListFile.ListCount - 1
          Form102.ListFile.ListIndex = Form102.pid
          End If
          Form102.MediaPlayer1.Filename = Form102.ListFile.List(Form102.pid)
           Form102.Label1.Caption = "media"
           Form102.TrackSelection.Left = 10000
           Form102.Frame8.Left = 2880
           Form102.Frame1.Left = 10000

        Unload Me
  End If
  Else
  If Form102.FileExists(Label3.Caption + "\SmM_FP.exe") = True Then
         Shell Label3.Caption + "\SmM_FP.exe", vbNormalFocus
       Else
         MsgBox ("找不到文件[ " + Label3.Caption + "\SmM_FP.exe" + " ].该文件可能已经丢失或被移动,请重新安装 Snowman Media ilxz 3.5")
       End If
         Unload Me
   End If
End Sub
Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label13.BackColor = &HFF0000
Picture9.BackColor = &HFF0000
Label13.ForeColor = &HFFFF&
End Sub
Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label14.BackColor = &HFF0000
Picture10.BackColor = &HFF0000
Label14.ForeColor = &HFFFF&
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5.BackColor = &HFF0000
Picture1.BackColor = &HFF0000
Label5.ForeColor = &HFFFF&
End Sub
