VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3630
   ClientLeft      =   6150
   ClientTop       =   165
   ClientWidth     =   4500
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3630
   ScaleWidth      =   4500
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   0
      Picture         =   "Form2.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4500
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   4020
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4470
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
      DisplayBackColor=   0
      DisplayForeColor=   16777215
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
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
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
Attribute VB_Name = "Form2"
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
Private Sub Form_Load()
 ingreturnvalue = SetWindowPos(Me.hwnd, HWND_TOPMOST, Val(10), Val(10), Val(10), Val(10), SWP_SHOWWINDOW)
    Form2.Width = 4620
    Form2.Height = 4035


Form2.MediaPlayer1.filename = Form1.MediaPlayer1.filename
Form2.Caption = Form1.MediaPlayer1.filename & " - Snowman Media  Video Window"
Form1.MediaPlayer1.filename = "Open.wav"
'Form2.Image1.Height = Me.ScaleHeight - 660
'Form2.Image1.Width = Me.ScaleWidth
End Sub
Private Sub form_resize()
Form2.MediaPlayer1.Height = Form2.Height
Form2.MediaPlayer1.Width = Me.ScaleWidth
'Form2.Image1.Height = Form2.Height - 400
'Form2.Image1.Width = Me.ScaleWidth
Image1.Top = Form2.Height / 2 - Form2.Image1.Height / 2 - 350
Image1.Left = Form2.Width / 2 - Form2.Image1.Width / 2 - 80
End Sub

Private Sub Image1_Click()

End Sub
