VERSION 5.00
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "SmM_Tools.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "媒体向导"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5595
   Icon            =   "向导.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   5595
   StartUpPosition =   2  '屏幕中心
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3915
      TabIndex        =   1
      Top             =   4905
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   2  'Horizontal Line
      ForeColor       =   &H00000000&
      Height          =   4500
      Left            =   0
      Picture         =   "向导.frx":1AFA
      ScaleHeight     =   4500
      ScaleWidth      =   5625
      TabIndex        =   0
      Top             =   0
      Width           =   5630
      Begin VB.Image Image6 
         Appearance      =   0  'Flat
         Height          =   780
         Left            =   4050
         MouseIcon       =   "向导.frx":AB66
         MousePointer    =   99  'Custom
         Picture         =   "向导.frx":ACB8
         Stretch         =   -1  'True
         Top             =   90
         Width           =   870
      End
      Begin VB.Image Image5 
         Appearance      =   0  'Flat
         Height          =   960
         Left            =   1395
         MouseIcon       =   "向导.frx":ACEF
         MousePointer    =   99  'Custom
         Picture         =   "向导.frx":AE41
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1005
      End
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         Height          =   1230
         Left            =   135
         MouseIcon       =   "向导.frx":AE78
         MousePointer    =   99  'Custom
         Picture         =   "向导.frx":AFCA
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1185
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   1050
         Left            =   4545
         MouseIcon       =   "向导.frx":B001
         MousePointer    =   99  'Custom
         Picture         =   "向导.frx":B153
         Stretch         =   -1  'True
         Top             =   1890
         Width           =   960
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   915
         Left            =   3285
         MouseIcon       =   "向导.frx":B18A
         MousePointer    =   99  'Custom
         Picture         =   "向导.frx":B2DC
         Stretch         =   -1  'True
         Top             =   3465
         Width           =   915
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   630
         MouseIcon       =   "向导.frx":B313
         MousePointer    =   99  'Custom
         Picture         =   "向导.frx":B465
         Stretch         =   -1  'True
         Top             =   2700
         Width           =   1590
      End
      Begin VB.Image Image7 
         Appearance      =   0  'Flat
         Height          =   1380
         Left            =   550
         Top             =   1255
         Width           =   3525
      End
   End
   Begin API控制大全.LyfTools Lf 
      Left            =   4860
      Top             =   4860
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6390
      Top             =   4815
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu dsfcvx 
      Caption         =   "a"
      Visible         =   0   'False
      Begin VB.Menu dsfsdc 
         Caption         =   "音频 CD (&A)"
      End
      Begin VB.Menu asfrgvvvv 
         Caption         =   "-"
      End
      Begin VB.Menu cdscew 
         Caption         =   "VCD (&V)"
      End
      Begin VB.Menu ceecfdf 
         Caption         =   "-"
      End
      Begin VB.Menu sdfe 
         Caption         =   "DVD (&D)"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A01 As Boolean
Dim A02 As Boolean
Dim A03 As Boolean
Dim A04 As Boolean
Dim A05 As Boolean
Dim A06 As Boolean
Dim CDRom As String
Private Sub cdscew_Click()
Dim i As Integer
If Lf.FileExists(CDRom + "MPEGAV\AVSEQ01.DAT") = True Or Lf.FileExists(CDRom + "MPEGAV\MUSIC01.DAT") = True Then
            File1.Path = CDRom + "MPEGAV"
          Open App.Path + "\SmM_List.sml" For Output As #1
    For i = 0 To File1.ListCount - 1
     Print #1, CDRom + "MPEGAV\" + File1.List(i)
    Next i
   Close (1)
            Shell (App.Path + "\Snowman Media ilxz.exe " + App.Path + "\SmM_List.sml")
Unload Me
Else
MsgBox "找不到 VCD 视频光盘,请重新插入。", vbExclamation
End If
End Sub


Private Sub dsfsdc_Click()
Dim i As Integer
If Lf.FileExists(CDRom + "Track01.cda") = True Then
File1.Path = CDRom
Open App.Path + "\SmM_List.sml" For Output As #1
    For i = 0 To File1.ListCount - 1
     Print #1, CDRom + File1.List(i)
    Next i
   Close (1)
Shell (App.Path + "\Snowman Media ilxz.exe " + App.Path + "\SmM_List.sml")
Unload Me
Else
MsgBox "找不到音频 CD 唱片,请重新插入。", vbExclamation

End If
End Sub

Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance Then End
'Lf.MakeTop Me, True
'Sf1.SkinPath = App.Path + "\SmM_Skin"
Dim DriveType As Integer
Dim rtn As String
Dim AllDrives As String
Dim JustOneDrive As String
AllDrives = Space$(64)

rtn = GetLogicalDriveStrings(Len(AllDrives), AllDrives) 'call the function to get the string containing all drives
AllDrives = Left(AllDrives, rtn) 'trim off trailing chr(0)'s.  AllDrives$ now contains all the drive letters.

Do
  rtn = InStr(AllDrives, Chr(0)) 'find the first separating chr(0)
  If rtn Then 'if there is one then
     JustOneDrive = Left(AllDrives, rtn) 'extract the drive up to the chr(0)
     AllDrives = Mid(AllDrives, rtn + 1, Len(AllDrives)) 'and remove that from the Alldrives string, so it won't be checked again
     
     rtn = GetDriveType(JustOneDrive) 'check what drive it is
     If rtn = DRIVE_CDROM Then 'if it is a CD-Rom drive then
        CDRom = Left(UCase(JustOneDrive), 3) 'return the drive letter to the user
        Exit Do
     End If
  End If
Loop Until AllDrives = "" Or DriveType = DRIVE_CDROM

End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If A01 = False Then
Image7.Picture = LoadPicture(App.Path + "\SmM_Start\003.gif")
A01 = True
End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
          CommonDialog1.Filter = "媒体文件 (多种被支持的类型)" & _
          "|*.smm;*.sma;*.smv;*.sml;*.ilxz;*.asf;*.asx;*.wm;*.wmx;*.wmp;*.wma;*.wax;*.wmv;*.wvx;*.vob;*.cda;*.wav;*.avi;*.mpeg;*.mpg;*.mpe;*.m1v;*.mp2;*.mpv2;*.mp2v;*.mpa;*.mp3;*.m3u;*.mid;*.midi;*.rmi;*.ivf;*.aif;*.aifc;*.aiff;*.au;*.snd;*.swf|图片文件 (*.bmp;*.jpg;*.gif)|*.bmp;*jpg;*.gif|所有文件 (*.*)|*.*"
          CommonDialog1.FilterIndex = 1

CommonDialog1.FileName = ""
CommonDialog1.DialogTitle = "浏览要播放的媒体"
          CommonDialog1.ShowOpen
          If Len(CommonDialog1.FileName) > 0 Then
     Shell App.Path + "\Snowman Media ilxz.exe " + CommonDialog1.FileName, vbNormalFocus
 Unload Me
     End If
End If
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If A02 = False Then
Image7.Picture = LoadPicture(App.Path + "\SmM_Start\007.gif")
A02 = True
End If
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "尚未安装 Snowman Media 网络收音机调谐器插件。", vbExclamation
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If A03 = False Then
Image7.Picture = LoadPicture(App.Path + "\SmM_Start\006.gif")
A03 = True
End If
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shell App.Path + "\SmM_HDTV.exe", vbNormalFocus
Unload Me
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If A04 = False Then
Image7.Picture = LoadPicture(App.Path + "\SmM_Start\002.gif")
A04 = True
End If
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PopupMenu dsfcvx, 0, X + 135, Y + 720

End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If A05 = False Then
Image7.Picture = LoadPicture(App.Path + "\SmM_Start\004.gif")
A05 = True
End If
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
          CommonDialog1.Filter = "Flash 影片 (多种被支持的类型)" & _
          "|*.swf|所有文件 (*.*)|*.*"
          CommonDialog1.FilterIndex = 1

CommonDialog1.FileName = ""
CommonDialog1.DialogTitle = "浏览要播放的 Flash 影片"
          CommonDialog1.ShowOpen
          If Len(CommonDialog1.FileName) > 0 Then
     Shell App.Path + "\SmM_Flash.exe " + CommonDialog1.FileName, vbNormalFocus
 Unload Me
     End If
End If

End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If A06 = False Then
Image7.Picture = LoadPicture(App.Path + "\SmM_Start\005.gif")
A06 = True
End If
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim SelectFileName As String
aa:
 SelectFileName = InputBox("请输入万维网地址 (URL) 或指定你要打开的本地文件路径。", , SelectFileName)
  If Len(SelectFileName) > 0 Then
 If Lf.FileExists(SelectFileName) = True Then
      Shell App.Path + "\Snowman Media ilxz.exe " + SelectFileName
      Unload Me
 Else
  If MsgBox("所请求的媒体文件不存在,如果是万维网资源请先连接网络。", vbRetryCancel) = vbRetry Then
     GoTo aa:
   Else
   Unload Me
   End If
   End If
     End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If A01 = True Then
Image7.Picture = LoadPicture("")
A01 = False
End If
If A02 = True Then
Image7.Picture = LoadPicture("")
A02 = False
End If
If A03 = True Then
Image7.Picture = LoadPicture("")
A03 = False
End If
If A04 = True Then
Image7.Picture = LoadPicture("")
A04 = False
End If
If A05 = True Then
Image7.Picture = LoadPicture("")
A05 = False
End If
If A06 = True Then
Image7.Picture = LoadPicture("")
A06 = False
End If

End Sub

Private Sub sdfe_Click()
Shell App.Path + "\SmM_DVD.exe", vbNormalFocus
Unload Me
End Sub
