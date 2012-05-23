VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{549BB3E0-192D-11D4-881C-F67CE1A5B67B}#1.0#0"; "MP3ENCODEDEMO.OCX"
Object = "{3B00B10A-6EF0-11D1-A6AA-0020AFE4DE54}#1.0#0"; "MP3PLAY.OCX"
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "LYFTOOLS.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form4 
   Caption         =   "Snowman Media Digtial Audio Editor  2.0"
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7740
   Icon            =   "数字音频.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7740
   StartUpPosition =   2  '屏幕中心
   Begin MP3Encode20.MP3Encode MP3Encode1 
      Height          =   390
      Left            =   1755
      TabIndex        =   4
      Top             =   7695
      Width           =   390
      Object.Visible         =   -1  'True
      Mode            =   1
      BladePath       =   ""
      Bitrate         =   0
      InputFile       =   ""
      Priority        =   0
   End
   Begin MPEGPLAYLib.Mp3Play MP 
      Height          =   375
      Left            =   1170
      TabIndex        =   3
      Top             =   7560
      Width           =   420
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6165
      Top             =   6660
   End
   Begin API控制大全.LyfTools LyfTools1 
      Left            =   3195
      Top             =   6795
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin ACTIVESKINLibCtl.SkinForm SkinForm1 
      Height          =   480
      Left            =   2340
      OleObjectBlob   =   "数字音频.frx":1582
      TabIndex        =   5
      Top             =   6930
      Width           =   480
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6315
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   11139
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "资源文件"
      TabPicture(0)   =   "数字音频.frx":15CB
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "目标文件"
      TabPicture(1)   =   "数字音频.frx":15E7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "Frame4"
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame8 
         Caption         =   "目标文件:"
         Height          =   4920
         Left            =   -74820
         TabIndex        =   43
         Top             =   450
         Width           =   4110
         Begin VB.DriveListBox Drive2 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   180
            TabIndex        =   46
            Top             =   270
            Width           =   3750
         End
         Begin VB.DirListBox Dir2 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000012&
            Height          =   3450
            Left            =   180
            TabIndex        =   45
            Top             =   630
            Width           =   3750
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000012&
            Height          =   285
            Left            =   180
            TabIndex        =   44
            Text            =   "Filename"
            Top             =   4500
            Width           =   3750
         End
         Begin VB.Label Label2 
            Caption         =   "文件名(&N):"
            Height          =   240
            Left            =   180
            TabIndex        =   47
            Top             =   4230
            Width           =   1185
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "目标文件:"
         Height          =   4245
         Left            =   180
         TabIndex        =   38
         Top             =   450
         Width           =   5595
         Begin VB.DirListBox Dir1 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000012&
            Height          =   2820
            Left            =   180
            TabIndex        =   41
            Top             =   720
            Width           =   2310
         End
         Begin VB.FileListBox File1 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000012&
            Height          =   3330
            Left            =   2655
            TabIndex        =   40
            Top             =   225
            Width           =   2760
         End
         Begin VB.DriveListBox Drive1 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   180
            TabIndex        =   39
            Top             =   270
            Width           =   2310
         End
         Begin VB.Label MName 
            Caption         =   "曲目     :"
            Height          =   390
            Left            =   180
            TabIndex        =   42
            Top             =   3690
            Width           =   5205
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "控制按钮:"
         Height          =   735
         Left            =   -74820
         TabIndex        =   31
         Top             =   5445
         Width           =   7395
         Begin VB.CommandButton Command1 
            Height          =   330
            Index           =   4
            Left            =   3060
            Picture         =   "数字音频.frx":1603
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "转换到WAV文件"
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   465
         End
         Begin VB.CommandButton Command2 
            Caption         =   "●"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2610
            TabIndex        =   49
            Top             =   270
            Width           =   465
         End
         Begin VB.CommandButton Command1 
            Height          =   330
            Index           =   2
            Left            =   1080
            Picture         =   "数字音频.frx":1B45
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "停止"
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   465
         End
         Begin VB.CommandButton Command1 
            Height          =   330
            Index           =   7
            Left            =   2070
            Picture         =   "数字音频.frx":2087
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "前进100帧"
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   465
         End
         Begin VB.CommandButton Command1 
            Height          =   330
            Index           =   8
            Left            =   1620
            Picture         =   "数字音频.frx":25C9
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "倒退100帧"
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   465
         End
         Begin VB.CommandButton Command1 
            Height          =   330
            Index           =   1
            Left            =   630
            Picture         =   "数字音频.frx":2B0B
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "暂停"
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   465
         End
         Begin VB.CommandButton Command1 
            Height          =   330
            Index           =   0
            Left            =   180
            Picture         =   "数字音频.frx":304D
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "播放键"
            Top             =   270
            UseMaskColor    =   -1  'True
            Width           =   465
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mp3信息:"
         Height          =   1410
         Left            =   180
         TabIndex        =   26
         Top             =   4770
         Width           =   5595
         Begin VB.Label BitS 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "比特率   :"
            Height          =   180
            Left            =   180
            TabIndex        =   30
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Caiyin 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "采样频率 :"
            Height          =   180
            Left            =   180
            TabIndex        =   29
            Top             =   810
            Width           =   900
         End
         Begin VB.Label Zheng 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "长度     :"
            Height          =   180
            Left            =   180
            TabIndex        =   28
            Top             =   270
            Width           =   900
         End
         Begin VB.Label ShDao 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "立体声   :"
            Height          =   180
            Left            =   180
            TabIndex        =   27
            Top             =   540
            Width           =   900
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "模式:"
         Height          =   5730
         Left            =   5850
         TabIndex        =   22
         Top             =   450
         Width           =   1725
         Begin VB.OptionButton Option4 
            Caption         =   "Mp3 to Wav"
            Height          =   285
            Left            =   135
            TabIndex        =   25
            Top             =   675
            Width           =   1365
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Wav to Mp3"
            Height          =   285
            Left            =   135
            TabIndex        =   24
            Top             =   315
            Width           =   1410
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Others"
            Height          =   240
            Left            =   135
            TabIndex        =   23
            Top             =   1080
            Width           =   1140
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Wav to Mp3:"
         Height          =   2160
         Left            =   -70635
         TabIndex        =   14
         Top             =   450
         Width           =   3255
         Begin VB.OptionButton Option2 
            Caption         =   "立体声"
            Height          =   195
            Left            =   1530
            TabIndex        =   15
            Top             =   270
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Left            =   1215
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   900
            Width           =   1905
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Left            =   1215
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   540
            Width           =   1905
         End
         Begin VB.OptionButton Option1 
            Caption         =   "单声道"
            Height          =   192
            Left            =   180
            TabIndex        =   16
            Top             =   270
            Width           =   1305
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "优先级(&P):"
            Height          =   180
            Index           =   2
            Left            =   180
            TabIndex        =   21
            Top             =   945
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "比特率(&B):"
            Height          =   180
            Left            =   180
            TabIndex        =   20
            Top             =   585
            Width           =   900
         End
         Begin VB.Label Label6 
            Caption         =   "完成度   :   0%"
            Height          =   195
            Left            =   180
            TabIndex        =   19
            Top             =   1395
            Width           =   1725
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Mp3 to Wav:"
         Height          =   2715
         Left            =   -70635
         TabIndex        =   7
         Top             =   2655
         Width           =   3210
         Begin ComctlLib.Slider Sder 
            Height          =   300
            Left            =   135
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   "拖放滑动块可快进或回退"
            Top             =   1800
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   529
            _Version        =   327682
            Max             =   500
            SelectRange     =   -1  'True
            TickStyle       =   3
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1215
            TabIndex        =   9
            Text            =   "0"
            Top             =   630
            Width           =   1860
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Left            =   1215
            TabIndex        =   8
            Text            =   "0"
            Top             =   270
            Width           =   1860
         End
         Begin VB.Label Zhongchang 
            AutoSize        =   -1  'True
            Caption         =   "总长度   :"
            Height          =   180
            Left            =   180
            TabIndex        =   13
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "结束帧(&E):"
            Height          =   180
            Index           =   1
            Left            =   180
            TabIndex        =   12
            Top             =   675
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "起始帧(&S):"
            Height          =   180
            Index           =   0
            Left            =   180
            TabIndex        =   11
            Top             =   315
            Width           =   900
         End
         Begin VB.Label DangQian 
            AutoSize        =   -1  'True
            Caption         =   "当前帧   :"
            Height          =   180
            Left            =   180
            TabIndex        =   10
            Top             =   1440
            Width           =   900
         End
      End
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   1005
      Left            =   1170
      TabIndex        =   2
      Top             =   7695
      Width           =   2535
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
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
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   285
      Left            =   6345
      TabIndex        =   1
      Top             =   7605
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   420
      Left            =   4680
      TabIndex        =   0
      Top             =   7740
      Width           =   1140
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tk, oldd As Integer
Dim Pok As Long
Dim Pn, ph As String
Dim PKey As String
Dim Sheng As String
Dim ToWavSong As String
Dim PS As Long
Dim Comline As String

Private Sub Command1_Click(Index As Integer)
On Error Resume Next
Select Case Index
           Case 0
                If Len(MediaPlayer1.FileName) > 0 Then
                 MediaPlayer1.Play
                 Else:
                         If Len(Dir1.Path) = 3 Then
                         MediaPlayer1.FileName = Dir1.Path + File1.FileName
                         Else
                           MediaPlayer1.FileName = Dir1.Path + "\" + File1.FileName
                           End If
                 End If
           Case 1
                 MediaPlayer1.Pause
           Case 2
                MediaPlayer1.FileName = "LLXX"
           Case 4
                 If Label4.Caption = "mtw" Then
                      MP.Close
                      MPinfo
                      If Len(Dir2.Path) = 3 Then
                      ToWavSong = Dir2.Path + Text3.Text & ".wav"
                      Else
                    ToWavSong = Dir2.Path + "\" + Text3.Text & ".wav"
                      End If
                     If Len(Dir2.Path) = 3 Then
                     PS = MP.Open(Dir1.Path + File1.FileName, ToWavSong)
                    Else
                      PS = MP.Open(Dir1.Path + "\" + File1.FileName, ToWavSong)
                     End If
                      Sder.Value = Text1.Text
                      MP.Play
                      MP.Seek (Text1.Text)

                End If
                 If Label4.Caption = "wtm" Then
                 If Len(Dir1.Path) = 3 Then
                 MP3Encode1.InputFile = Dir1.Path + File1.FileName
                    Else: MP3Encode1.InputFile = Dir1.Path + "\" + File1.FileName
                 End If
                      MP3Encode1.Bitrate = Combo1.ListIndex
                      MP3Encode1.Priority = Combo2.ListIndex
                      If Option1.Value = True Then MP3Encode1.Mode = vbMono
                      If Option2.Value = True Then MP3Encode1.Mode = vbStereo
                      MP3Encode1.Encode

                 End If



End Select
End Sub
Sub MPinfo()
If Len(Dir1.Path) = 3 Then
MName.Caption = "曲目     : " & Left(Dir1.Path, 2) + "\" + File1.FileName
Else: MName.Caption = "曲目     : " & Dir1.Path + "\" + File1.FileName
End If
If Label4.Caption = "mtw" Then
MP.FrameNotifyCount = 10
Sder.Max = MP.FrameCount
If MP.MPEGMode = 4 Then
        Sheng = "立体声   : Х"
ElseIf MP.MPEGMode = 2 Then
        Sheng = "立体声   : √"
Else: Sheng = "立体声   :"
End If
ShDao.Caption = Sheng
Zheng.Caption = "长度     : " & MP.FrameCount & "帧  合计: " & Format(MP.FrameCount * MP.MsPerFrame \ 1000 \ 60, "#0 分 ") & Format((MP.FrameCount * MP.MsPerFrame) \ 1000 Mod 60, "00 秒")
Caiyin.Caption = "采样频率 : " & MP.SampleFrequency \ 1000 & " kHz"
BitS.Caption = "比特率   : " & MP.Bitrate \ 1000 & " kbps"
Text2.Text = MP.FrameCount
ZHinfo
End If
End Sub
Sub ZHinfo()
On Error Resume Next
Zhongchang.Caption = "总长度   :  " & Format((Text2.Text - Text1.Text) * MP.MsPerFrame / 1000, "#0.00 秒")
Sder.SelStart = Text1.Text
Sder.SelLength = Text2.Text - Text1.Text
oldd = MP.MsPerFrame
End Sub

Private Sub Command2_Click()
If Label4.Caption = "wtm" Then
PS = MP.Open(PlaySong, "")
Sder.Value = Text1.Text
MP.Play
MP.Seek (Text1.Text)
End If
End Sub

Private Sub Drive2_Change()
Dir2.Path = Drive2
End Sub

Private Sub File1_Click()
        MP.Close
        If Len(Dir1.Path) = 3 Then
        PS = MP.Open(Left(Dir1.Path, 2) + "\" + File1.FileName, "")
       Else: PS = MP.Open(Dir1.Path + "\" + File1.FileName, "")
       End If
        MPinfo
End Sub

Private Sub MP_ActFrame(ByVal ActFrame As Long)
Sder.Value = ActFrame
End Sub
Private Sub Sder_Change()
On Error Resume Next
DangQian.Caption = "当前帧   :  " & Sder.Value & " - " & Format(Sder.Value * MP.MsPerFrame / 1000, "#0.00s")
End Sub
Private Sub Text1_Change()
ZHinfo
End Sub
Private Sub Text2_Change()
ZHinfo
End Sub
Private Sub MP3Encode1_OnFinished()
 MsgBox ("已经完成指定曲目的转换.")
 Beep
End Sub
Private Sub timer1_timer()
Label6.Caption = "完成度   :  " & CStr(MP3Encode1.Position) & "%"
If Sder.Value > Text2.Text - 30 Or Sder = Text2.Text Then
MP.Close
End If
End Sub
Private Sub Dir1_Change()
File1.Path = Dir1
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1
End Sub
Private Sub File1_DblClick()
If Len(Dir1.Path) = 3 Then
MediaPlayer1.FileName = Left(Dir1.Path, 2) + "\" + File1.FileName
Else
MediaPlayer1.FileName = Dir1.Path + "\" + File1.FileName
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
SkinForm1.SkinPath = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Skin_Path")
Label5.Caption = App.Path
Combo1.AddItem (32)
Combo1.AddItem (40)
Combo1.AddItem (48)
Combo1.AddItem (56)
Combo1.AddItem (64)
Combo1.AddItem (80)
Combo1.AddItem (96)
Combo1.AddItem (112)
Combo1.AddItem (128)
Combo1.AddItem (160)
Combo1.AddItem (192)
Combo1.AddItem (224)
Combo1.AddItem (256)
Combo1.AddItem (320)
Combo1.ListIndex = 8
Combo2.AddItem ("较低")
Combo2.AddItem ("标准")
Combo2.AddItem ("较高")
Combo2.ListIndex = 1
Option5.Value = False
Option4.Value = False
Option3.Value = True
File1.Pattern = "*.wav"
Label4.Caption = "wtm"
Pn = "ufo2000"
PKey = "1208686909"
Pok = MP.Authorize(Pn, PKey)
MP3Encode1.BladePath = Label5.Caption + "\bladeenc.exe"
End Sub
Private Sub option3_Click()
Label4.Caption = "wtm"
File1.Pattern = "*.wav"
Frame1.Enabled = False
Frame4.Enabled = False
ShDao.Enabled = False
DangQian.Enabled = False
Zhongchang.Enabled = False
Caiyin.Enabled = False
BitS.Enabled = False
Zheng.Enabled = False
Label1(0).Enabled = False
Label1(1).Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Frame3.Enabled = True
Label3.Enabled = True
Label1(2).Enabled = True
Label6.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Command1(4).Enabled = True
End Sub
Private Sub option4_Click()
Label4.Caption = "mtw"
File1.Pattern = "*.mp3"
Command1(4).Enabled = True
Frame3.Enabled = False
Label3.Enabled = False
Label1(2).Enabled = False
Label6.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Frame1.Enabled = True
Frame4.Enabled = True
ShDao.Enabled = True
DangQian.Enabled = True
Zhongchang.Enabled = True
Caiyin.Enabled = True
BitS.Enabled = True
Zheng.Enabled = True
Label1(0).Enabled = True
Label1(1).Enabled = True
Text1.Enabled = True
Text2.Enabled = True
End Sub
Private Sub option5_Click()
File1.Pattern = "*.*"
End Sub
