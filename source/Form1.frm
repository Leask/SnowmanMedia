VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   3000
   ClientTop       =   0
   ClientWidth     =   5235
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   3375
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   855
      Top             =   855
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Snowman Media  Files Opening Window"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Snowman Media"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   225
      TabIndex        =   1
      ToolTipText     =   "显示当前播放媒体的路径和文件名  －－Snowman Media"
      Top             =   0
      Width           =   3480
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "http://www.h2ont.com"
         Height          =   195
         Left            =   270
         TabIndex        =   2
         Top             =   180
         Width           =   1815
      End
   End
   Begin VB.Image Image17 
      Height          =   375
      Left            =   1215
      MouseIcon       =   "Form1.frx":0B1A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":0C6C
      Stretch         =   -1  'True
      ToolTipText     =   "轻松播放CD和VCD  －－Snowman Media"
      Top             =   2295
      Width           =   660
   End
   Begin VB.Image Image2 
      Height          =   1695
      Left            =   45
      Picture         =   "Form1.frx":106B
      ToolTipText     =   "播放视频的窗口  －－Snowman Media"
      Top             =   405
      Width           =   2115
   End
   Begin VB.Image Image16 
      Height          =   1290
      Left            =   1080
      Picture         =   "Form1.frx":306E
      ToolTipText     =   "播放状态提示栏  －－Snowman Media"
      Top             =   2295
      Width           =   960
   End
   Begin VB.Image Image12 
      Height          =   420
      Left            =   2475
      MouseIcon       =   "Form1.frx":3746
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":3898
      Stretch         =   -1  'True
      ToolTipText     =   "在线播放多种网络媒体  －－Snowman Media"
      Top             =   1305
      Width           =   2325
   End
   Begin VB.Image Image15 
      Height          =   510
      Left            =   2475
      MouseIcon       =   "Form1.frx":3C97
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":3DE9
      Stretch         =   -1  'True
      ToolTipText     =   "打开和编辑自己的媒体播放列表  －－Snowman Media"
      Top             =   1710
      Width           =   2325
   End
   Begin VB.Image Image14 
      Height          =   465
      Left            =   2070
      MouseIcon       =   "Form1.frx":41E8
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":433A
      Stretch         =   -1  'True
      ToolTipText     =   "显示在线帮助  －－Snowman Media"
      Top             =   2655
      Width           =   2280
   End
   Begin VB.Image Image13 
      Height          =   420
      Left            =   2340
      MouseIcon       =   "Form1.frx":4739
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":488B
      Stretch         =   -1  'True
      ToolTipText     =   "打开Flash动画文件  －－Snowman Media"
      Top             =   2205
      Width           =   2325
   End
   Begin VB.Image Image11 
      Height          =   465
      Left            =   2385
      MouseIcon       =   "Form1.frx":4C8A
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":4DDC
      Stretch         =   -1  'True
      ToolTipText     =   "打开多种格式的图片文件  －－Snowman Media"
      Top             =   810
      Width           =   2280
   End
   Begin VB.Image Image10 
      Height          =   375
      Left            =   4725
      MouseIcon       =   "Form1.frx":51DB
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":532D
      Stretch         =   -1  'True
      ToolTipText     =   "关闭  Snowman Media"
      Top             =   0
      Width           =   300
   End
   Begin VB.Image Image9 
      Height          =   330
      Left            =   4410
      MouseIcon       =   "Form1.frx":572C
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":587E
      Stretch         =   -1  'True
      ToolTipText     =   "最小化  Snowman Media"
      Top             =   45
      Width           =   300
   End
   Begin VB.Image Image8 
      Height          =   330
      Left            =   4095
      MouseIcon       =   "Form1.frx":5C7D
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":5DCF
      Stretch         =   -1  'True
      ToolTipText     =   "关于  Snowman Media"
      Top             =   45
      Width           =   315
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   3780
      MouseIcon       =   "Form1.frx":61CE
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":6320
      Stretch         =   -1  'True
      ToolTipText     =   "打开视频拖放窗口  －－Snowman Media"
      Top             =   0
      Width           =   300
   End
   Begin VB.Image Image5 
      Height          =   750
      Left            =   0
      Picture         =   "Form1.frx":671F
      ToolTipText     =   "播放状态提示栏  －－Snowman Media"
      Top             =   2700
      Width           =   2040
   End
   Begin VB.Image Image7 
      Height          =   420
      Left            =   2160
      MouseIcon       =   "Form1.frx":6B68
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":6CBA
      Stretch         =   -1  'True
      ToolTipText     =   "打开播放30多种视频和音频的媒体文件  －－Snowman Media"
      Top             =   360
      Width           =   2280
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      DragIcon        =   "Form1.frx":70B9
      Height          =   2715
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   1995
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
      DisplayForeColor=   0
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
   Begin VB.Image Image1 
      Height          =   2700
      Left            =   1890
      Picture         =   "Form1.frx":720B
      Top             =   405
      Width           =   3345
   End
   Begin VB.Image Image3 
      Height          =   585
      Left            =   0
      Picture         =   "Form1.frx":A6AB
      Top             =   0
      Width           =   585
   End
   Begin VB.Image Image4 
      Height          =   585
      Left            =   3240
      Picture         =   "Form1.frx":AA3A
      Top             =   0
      Width           =   2010
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim num As Integer
Dim filename As String
Dim filenum As Integer
Dim i As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim intpos As Integer
Const hwnd_top = 0
Dim ingreturnvalue As Long
Const HWND_TOPMOST = -1
Const SWP_SHOWWINDOW = &H40

Private Sub Form_Unload(Cancel As Integer)
Dim z As Integer
  End
End Sub

Private Sub image17_click()
Form10.Show
End Sub
Private Sub image7_Click()
    Form1.CommonDialog1.Filter = "流行的三十多种媒体文件:cd、vcd、mp3、wma、wav、wax、asf、rmi、asx、mov、m1v、mp2、mpg、mpeg、mpa、mpe、avi、mid、qt、m3u、aif、aifc、aiff、au、snd..." & _
    "|*.au;*.dat;*.and;*.aif;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.m3u;*.wma;*.wav;*.avi;*.mid|浏览所有文件:*.*|*.*"
    
    Form1.CommonDialog1.FilterIndex = 1
    Form1.CommonDialog1.ShowOpen
    If Len(Form1.CommonDialog1.filename) > 0 Then
    Form1.MediaPlayer1.filename = Form1.CommonDialog1.filename
    Form1.Caption = Form1.CommonDialog1.filename & " - Snowman Media"
    Form1.Label1.Caption = Form1.CommonDialog1.filename
End If
End Sub
  
  Private Sub image10_click()
    Unload Me
End Sub
  
  Private Sub image13_click()
   Form1.CommonDialog1.Filter = "Flash文件(*.swf)" & _
    "|*.swf|浏览所有文件:*.*|*.*"
    
    Form1.CommonDialog1.FilterIndex = 1

    Form1.CommonDialog1.ShowOpen
    If Len(Form1.CommonDialog1.filename) > 0 Then
    Form7.ShockwaveFlash1.Movie = Form1.CommonDialog1.filename
    Form7.Caption = Form1.CommonDialog1.filename & " - Snowman Media"
    Form7.Show
   End If
  End Sub
  Private Sub image15_click()
  Form6.Show
  End Sub
  Private Sub image12_click()
  Form5.Show
  End Sub
  
  
  
  
  Private Sub image11_click()
  Form1.CommonDialog1.Filter = "流行的图片文件:*.bmp、*.jpg、*.did、*.wmf、*.ico、*.gif、*.rle、*.cur、*.emf" & _
    "|*.bmp;*.jpg;*.did;*.wmf;*.ico;*.gif;*.rle;*.cur;*.emf|浏览所有文件:*.*|*.*"
    
    Form1.CommonDialog1.FilterIndex = 1
    Form1.CommonDialog1.ShowOpen
   If Len(Form1.CommonDialog1.filename) > 0 Then
    Form8.Image1.Picture = LoadPicture(Form1.CommonDialog1.filename)
    Form8.Image1.Picture = LoadPicture(Form1.CommonDialog1.filename)
    Form8.Caption = Form1.CommonDialog1.filename & " - Snowman Media"
    Form8.Show
   End If
  End Sub

  
  Private Sub image9_click()
Me.Visible = False
    
    
    
    
    End Sub

  
  Private Sub image8_click()
  Form4.Show
  End Sub
  
  Private Sub image14_Click()
  Form3.Show
  End Sub
  
  Private Sub image6_click()
  
  Form2.Show
  
  
  End Sub
Private Sub form_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub form_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub

Private Sub Form_Load()
     Form9.Show

     ingreturnvalue = SetWindowPos(Me.hwnd, hwnd_top, Val(10), Val(10), Val(10), Val(10), SWP_SHOWWINDOW)
    Form1.Width = 5235
    Form1.Height = 3105
    Form1.Caption = "Snowman Media"
    Form1.MediaPlayer1.ToolTipText = "嘘,静下来慢慢欣赏！   －－思夏祝你有愉悦的多媒体体验"
 End Sub
 
Private Sub image1_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub

Private Sub image1_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub

Private Sub mediaplayer1_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub



Private Sub mediaplayer1_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub frame1_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub frame1_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image2_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image2_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image3_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image3_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image4_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image4_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image5_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image5_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub

Private Sub label1_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub label1_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image7_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image7_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image6_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image6_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image8_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image8_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image9_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image9_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image10_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image10_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image11_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image11_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image12_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image12_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image13_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image13_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image14_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image14_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image15_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image15_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image16_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image16_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub
Private Sub image17_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image17_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form1.Left = Form1.Left + (c - a)
Form1.Top = Form1.Top + (d - b)
End Sub

