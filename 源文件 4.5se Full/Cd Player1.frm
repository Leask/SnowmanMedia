VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ActiveSkin.ocx"
Object = "{244E6785-6684-11D2-943F-A976CFB4FC0C}#1.0#0"; "ctlstbar.ocx"
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "lyftools.ocx"
Object = "{C40E7B9F-6CF0-11D2-AA70-444553540000}#1.0#0"; "Coolineprj.ocx"
Object = "{CFCDAA00-8BE4-11CF-B84B-0020AFBBCCFA}#1.0#0"; "rmoc3260.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form102 
   Appearance      =   0  'Flat
   Caption         =   "Snowman Media ilxz 3.5"
   ClientHeight    =   7680
   ClientLeft      =   4200
   ClientTop       =   2610
   ClientWidth     =   7395
   ForeColor       =   &H00000000&
   Icon            =   "Cd Player1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7680
   ScaleWidth      =   7395
   StartUpPosition =   2  '屏幕中心
   Begin CooLinePrj.CooLine CooLine1 
      Height          =   285
      Left            =   45
      TabIndex        =   72
      ToolTipText     =   "信息"
      Top             =   45
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   503
      InsChr          =   95
      Display         =   "H2ont Leask Snowman Media ilxz 3.5 Plus Edition Ready Now！   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8454143
      ForeColor       =   16711680
   End
   Begin ACTIVESKINLibCtl.SkinForm SkinForm1 
      Height          =   480
      Left            =   1575
      OleObjectBlob   =   "Cd Player1.frx":1582
      TabIndex        =   65
      Top             =   7065
      Visible         =   0   'False
      Width           =   480
   End
   Begin API控制大全.LyfTools LyfTools1 
      Left            =   945
      Top             =   7110
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin CTLISTBARLibCtl.ctListBar ctListBar1 
      Height          =   5460
      Left            =   45
      TabIndex        =   20
      ToolTipText     =   "功能菜单"
      Top             =   405
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   9631
      _StockProps     =   70
      Caption         =   "正在播放"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackImage       =   "Cd Player1.frx":15CB
      ListBackColor   =   0
      ListForeColor   =   16711680
      BarForeColor    =   16711680
      WordWrap        =   -1  'True
      Caption         =   "正在播放"
      PicArray0       =   "Cd Player1.frx":15E7
      PicArray1       =   "Cd Player1.frx":1901
      PicArray2       =   "Cd Player1.frx":1C1B
      PicArray3       =   "Cd Player1.frx":206D
      PicArray4       =   "Cd Player1.frx":2387
      PicArray5       =   "Cd Player1.frx":26A1
      PicArray6       =   "Cd Player1.frx":29BB
      PicArray7       =   "Cd Player1.frx":2CD5
      PicArray8       =   "Cd Player1.frx":2FEF
      PicArray9       =   "Cd Player1.frx":38C9
      PicArray10      =   "Cd Player1.frx":3BE3
      PicArray11      =   "Cd Player1.frx":3EFD
      PicArray12      =   "Cd Player1.frx":4217
      PicArray13      =   "Cd Player1.frx":4531
      PicArray14      =   "Cd Player1.frx":484B
      PicArray15      =   "Cd Player1.frx":4B65
      PicArray16      =   "Cd Player1.frx":4FB7
      PicArray17      =   "Cd Player1.frx":52D1
      PicArray18      =   "Cd Player1.frx":55EB
      PicArray19      =   "Cd Player1.frx":5A3D
      PicArray20      =   "Cd Player1.frx":5D57
      PicArray21      =   "Cd Player1.frx":6071
      PicArray22      =   "Cd Player1.frx":638B
      PicArray23      =   "Cd Player1.frx":66A5
      PicArray24      =   "Cd Player1.frx":69BF
      PicArray25      =   "Cd Player1.frx":6CD9
      PicArray26      =   "Cd Player1.frx":6FF3
      PicArray27      =   "Cd Player1.frx":730D
      PicArray28      =   "Cd Player1.frx":858F
      PicArray29      =   "Cd Player1.frx":88A9
      PicArray30      =   "Cd Player1.frx":8CFB
      PicArray31      =   "Cd Player1.frx":9015
      PicArray32      =   "Cd Player1.frx":932F
      PicArray33      =   "Cd Player1.frx":9649
      PicArray34      =   "Cd Player1.frx":9963
      PicArray35      =   "Cd Player1.frx":9C7D
      PicArray36      =   "Cd Player1.frx":9F97
      PicArray37      =   "Cd Player1.frx":A2B1
      PicArray38      =   "Cd Player1.frx":A5CB
      PicArray39      =   "Cd Player1.frx":A8E5
      PicArray40      =   "Cd Player1.frx":ABFF
      PicArray41      =   "Cd Player1.frx":AC1B
      PicArray42      =   "Cd Player1.frx":AC37
      PicArray43      =   "Cd Player1.frx":AC53
      PicArray44      =   "Cd Player1.frx":AC6F
      PicArray45      =   "Cd Player1.frx":AC8B
      PicArray46      =   "Cd Player1.frx":ACA7
      PicArray47      =   "Cd Player1.frx":ACC3
      PicArray48      =   "Cd Player1.frx":ACDF
      PicArray49      =   "Cd Player1.frx":ACFB
      PicArray50      =   "Cd Player1.frx":AD17
      PicArray51      =   "Cd Player1.frx":AD33
      PicArray52      =   "Cd Player1.frx":AD4F
      PicArray53      =   "Cd Player1.frx":AD6B
      PicArray54      =   "Cd Player1.frx":AD87
      PicArray55      =   "Cd Player1.frx":ADA3
      PicArray56      =   "Cd Player1.frx":ADBF
      PicArray57      =   "Cd Player1.frx":ADDB
      PicArray58      =   "Cd Player1.frx":ADF7
      PicArray59      =   "Cd Player1.frx":AE13
      PicArray60      =   "Cd Player1.frx":AE2F
      PicArray61      =   "Cd Player1.frx":AE4B
      PicArray62      =   "Cd Player1.frx":AE67
      PicArray63      =   "Cd Player1.frx":AE83
      PicArray64      =   "Cd Player1.frx":AE9F
      PicArray65      =   "Cd Player1.frx":AEBB
      PicArray66      =   "Cd Player1.frx":AED7
      PicArray67      =   "Cd Player1.frx":AEF3
      PicArray68      =   "Cd Player1.frx":AF0F
      PicArray69      =   "Cd Player1.frx":AF2B
      PicArray70      =   "Cd Player1.frx":AF47
      PicArray71      =   "Cd Player1.frx":AF63
      PicArray72      =   "Cd Player1.frx":AF7F
      PicArray73      =   "Cd Player1.frx":AF9B
      PicArray74      =   "Cd Player1.frx":AFB7
      PicArray75      =   "Cd Player1.frx":AFD3
      PicArray76      =   "Cd Player1.frx":AFEF
      PicArray77      =   "Cd Player1.frx":B00B
      PicArray78      =   "Cd Player1.frx":B027
      PicArray79      =   "Cd Player1.frx":B043
      PicArray80      =   "Cd Player1.frx":B05F
      PicArray81      =   "Cd Player1.frx":B07B
      PicArray82      =   "Cd Player1.frx":B097
      PicArray83      =   "Cd Player1.frx":B0B3
      PicArray84      =   "Cd Player1.frx":B0CF
      PicArray85      =   "Cd Player1.frx":B0EB
      PicArray86      =   "Cd Player1.frx":B107
      PicArray87      =   "Cd Player1.frx":B123
      PicArray88      =   "Cd Player1.frx":B13F
      PicArray89      =   "Cd Player1.frx":B15B
      PicArray90      =   "Cd Player1.frx":B177
      PicArray91      =   "Cd Player1.frx":B193
      PicArray92      =   "Cd Player1.frx":B1AF
      PicArray93      =   "Cd Player1.frx":B1CB
      PicArray94      =   "Cd Player1.frx":B1E7
      PicArray95      =   "Cd Player1.frx":B203
      PicArray96      =   "Cd Player1.frx":B21F
      PicArray97      =   "Cd Player1.frx":B23B
      PicArray98      =   "Cd Player1.frx":B257
      PicArray99      =   "Cd Player1.frx":B273
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   1530
      TabIndex        =   12
      Top             =   4950
      Width           =   4515
      Begin VB.VScrollBar Volume 
         Height          =   915
         Left            =   4365
         MouseIcon       =   "Cd Player1.frx":B28F
         TabIndex        =   14
         Top             =   0
         Width           =   150
      End
      Begin VB.ListBox ListFile 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   930
         ItemData        =   "Cd Player1.frx":B3E1
         Left            =   0
         List            =   "Cd Player1.frx":B3E3
         OLEDropMode     =   1  'Manual
         TabIndex        =   13
         ToolTipText     =   "同步媒体列表"
         Top             =   0
         Width           =   4380
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   4470
      Left            =   1530
      TabIndex        =   3
      Top             =   405
      Width           =   4515
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   10000
         TabIndex        =   75
         ToolTipText     =   "状态"
         Top             =   4185
         Width           =   4425
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   135
            TabIndex        =   76
            ToolTipText     =   "状态"
            Top             =   5
            Width           =   90
         End
      End
      Begin RealAudioObjectsCtl.RealAudio RA 
         Height          =   3390
         Left            =   10000
         TabIndex        =   73
         Top             =   -23
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   5980
         AUTOSTART       =   -1  'True
         SHUFFLE         =   0   'False
         PREFETCH        =   0   'False
         NOLABELS        =   0   'False
         CONTROLS        =   "ImageWindow"
         LOOP            =   0   'False
         NUMLOOP         =   0
         CENTER          =   0   'False
         MAINTAINASPECT  =   0   'False
         BACKGROUNDCOLOR =   "#000000"
      End
      Begin VB.ComboBox TrackSelection 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   2655
         MouseIcon       =   "Cd Player1.frx":B3E5
         TabIndex        =   8
         Text            =   "S.M."
         ToolTipText     =   "曲目"
         Top             =   3735
         Width           =   780
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2745
         TabIndex        =   17
         ToolTipText     =   "方式"
         Top             =   3015
         Width           =   870
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "无序"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   0
            MaskColor       =   &H00FF0000&
            TabIndex        =   19
            ToolTipText     =   "无序方式"
            Top             =   0
            Width           =   690
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "循环"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   0
            MaskColor       =   &H00FF0000&
            TabIndex        =   18
            ToolTipText     =   "循环方式"
            Top             =   180
            Width           =   690
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   -270
         TabIndex        =   4
         ToolTipText     =   "状态"
         Top             =   2565
         Width           =   4380
         Begin VB.TextBox TimeWindow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   135
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "[00]00:00"
            ToolTipText     =   "播放时间"
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label TotalTrack 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   3060
            TabIndex        =   7
            ToolTipText     =   "总时间"
            Top             =   0
            Width           =   90
         End
         Begin VB.Label TrackTime 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   1800
            TabIndex        =   6
            ToolTipText     =   "曲目时间"
            Top             =   0
            Width           =   90
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   3690
         Width           =   5190
         Begin VB.Image Image16 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   3780
            Picture         =   "Cd Player1.frx":B537
            Stretch         =   -1  'True
            ToolTipText     =   "媒体播放"
            Top             =   0
            Width           =   345
         End
         Begin VB.Image Image17 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   4185
            Picture         =   "Cd Player1.frx":B936
            Stretch         =   -1  'True
            ToolTipText     =   "CD 播放"
            Top             =   0
            Width           =   435
         End
         Begin VB.Image Image5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   675
            MouseIcon       =   "Cd Player1.frx":BD35
            Picture         =   "Cd Player1.frx":BE87
            Stretch         =   -1  'True
            ToolTipText     =   "上一首曲目"
            Top             =   0
            Width           =   300
         End
         Begin VB.Image Image7 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1395
            MouseIcon       =   "Cd Player1.frx":C286
            Picture         =   "Cd Player1.frx":C3D8
            Stretch         =   -1  'True
            ToolTipText     =   "快进"
            Top             =   0
            Width           =   300
         End
         Begin VB.Image Image11 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2475
            MouseIcon       =   "Cd Player1.frx":C7D7
            Picture         =   "Cd Player1.frx":C929
            Stretch         =   -1  'True
            ToolTipText     =   "弹出光驱"
            Top             =   90
            Width           =   300
         End
         Begin VB.Image Image4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1755
            MouseIcon       =   "Cd Player1.frx":CD28
            Picture         =   "Cd Player1.frx":CE7A
            Stretch         =   -1  'True
            ToolTipText     =   "下一首曲目"
            Top             =   90
            Width           =   300
         End
         Begin VB.Image Image3 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   360
            MouseIcon       =   "Cd Player1.frx":D279
            Picture         =   "Cd Player1.frx":D3CB
            Stretch         =   -1  'True
            ToolTipText     =   "暂停"
            Top             =   135
            Width           =   255
         End
         Begin VB.Image Image18 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   0
            MouseIcon       =   "Cd Player1.frx":D7CA
            Picture         =   "Cd Player1.frx":D91C
            Stretch         =   -1  'True
            ToolTipText     =   "播放"
            Top             =   0
            Width           =   345
         End
         Begin VB.Image Image6 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1035
            MouseIcon       =   "Cd Player1.frx":DD1B
            Picture         =   "Cd Player1.frx":DE6D
            Stretch         =   -1  'True
            ToolTipText     =   "快退"
            Top             =   90
            Width           =   300
         End
         Begin VB.Image Image2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2115
            MouseIcon       =   "Cd Player1.frx":E26C
            Picture         =   "Cd Player1.frx":E3BE
            Stretch         =   -1  'True
            ToolTipText     =   "停止"
            Top             =   0
            Width           =   300
         End
         Begin VB.Image Image15 
            Appearance      =   0  'Flat
            Height          =   780
            Left            =   0
            Picture         =   "Cd Player1.frx":E7BD
            Top             =   0
            Width           =   4515
         End
      End
      Begin MSComctlLib.Slider SLD1 
         Height          =   330
         Left            =   10000
         TabIndex        =   74
         Top             =   3420
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   582
         _Version        =   393216
         LargeChange     =   20000
         SmallChange     =   20000
         Max             =   10000
         SelectRange     =   -1  'True
         TickStyle       =   3
      End
      Begin VB.Image Image10 
         Appearance      =   0  'Flat
         Height          =   2895
         Left            =   0
         OLEDropMode     =   1  'Manual
         Top             =   0
         Width           =   3705
      End
      Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
         DragIcon        =   "Cd Player1.frx":F82D
         Height          =   4470
         Left            =   45
         TabIndex        =   10
         Top             =   90
         Width           =   4515
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
         DisplayBackColor=   65535
         DisplayForeColor=   16711680
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
         ShowAudioControls=   -1  'True
         ShowDisplay     =   0   'False
         ShowGotoBar     =   0   'False
         ShowPositionControls=   -1  'True
         ShowStatusBar   =   -1  'True
         ShowTracker     =   -1  'True
         TransparentAtStart=   0   'False
         VideoBorderWidth=   0
         VideoBorderColor=   0
         VideoBorder3D   =   0   'False
         Volume          =   -600
         WindowlessVideo =   0   'False
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   2250
      Top             =   7155
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   7155
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Snowman Media  3.0"
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5460
      Left            =   90
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   450
      Width           =   1365
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      ForeColor       =   &H80000008&
      Height          =   4470
      Left            =   1575
      OLEDropMode     =   1  'Manual
      TabIndex        =   11
      Top             =   450
      Width           =   4515
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   1575
      OLEDropMode     =   1  'Manual
      TabIndex        =   15
      Top             =   4995
      Width           =   4515
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      ForeColor       =   &H80000008&
      Height          =   5460
      Left            =   6705
      TabIndex        =   21
      Top             =   405
      Width           =   4515
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2970
         ScaleHeight     =   255
         ScaleWidth      =   1470
         TabIndex        =   23
         Top             =   0
         Width           =   1500
         Begin VB.PictureBox Picture12 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2295
            ScaleHeight     =   255
            ScaleWidth      =   1065
            TabIndex        =   30
            Top             =   0
            Width           =   1095
            Begin VB.Label Label15 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "清除(&L)"
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   45
               TabIndex        =   31
               Top             =   135
               Width           =   1680
            End
         End
         Begin VB.PictureBox Picture14 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   1065
            TabIndex        =   28
            Top             =   4320
            Width           =   1095
            Begin VB.Label Label17 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "取消(&C)"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   45
               TabIndex        =   29
               Top             =   45
               Width           =   1635
            End
         End
         Begin VB.PictureBox Picture16 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   45
            ScaleHeight     =   255
            ScaleWidth      =   1335
            TabIndex        =   26
            Top             =   270
            Width           =   1365
            Begin VB.Label Label19 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "当前曲目(&T)"
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   45
               TabIndex        =   27
               Top             =   45
               Width           =   1680
            End
         End
         Begin VB.PictureBox Picture17 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   45
            ScaleHeight     =   255
            ScaleWidth      =   1335
            TabIndex        =   24
            Top             =   540
            Width           =   1365
            Begin VB.Label Label20 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "播放列表(&L)"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   45
               TabIndex        =   25
               Top             =   45
               Width           =   1635
            End
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            Caption         =   "预览(&P)"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   90
            TabIndex        =   32
            Top             =   45
            Width           =   1680
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1485
         ScaleHeight     =   255
         ScaleWidth      =   1470
         TabIndex        =   49
         Top             =   0
         Width           =   1500
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   45
            ScaleHeight     =   255
            ScaleWidth      =   1335
            TabIndex        =   56
            Top             =   270
            Width           =   1365
            Begin VB.Label Label12 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "添加(&T)"
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   45
               TabIndex        =   57
               Top             =   45
               Width           =   1680
            End
         End
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   45
            ScaleHeight     =   255
            ScaleWidth      =   1335
            TabIndex        =   54
            Top             =   540
            Width           =   1365
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "删除(&D)"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   45
               TabIndex        =   55
               Top             =   45
               Width           =   1635
            End
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   45
            ScaleHeight     =   255
            ScaleWidth      =   1335
            TabIndex        =   52
            Top             =   1125
            Width           =   1365
            Begin VB.Label Label10 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "清除(&L)"
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   45
               TabIndex        =   53
               Top             =   45
               Width           =   1680
            End
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   45
            ScaleHeight     =   255
            ScaleWidth      =   1335
            TabIndex        =   50
            Top             =   855
            Width           =   1365
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "全选(&A)"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   45
               TabIndex        =   51
               Top             =   45
               Width           =   1635
            End
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            Caption         =   "编辑(&E)"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   90
            TabIndex        =   58
            Top             =   45
            Width           =   1680
         End
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   1470
         TabIndex        =   37
         Top             =   0
         Width           =   1500
         Begin VB.PictureBox Picture15 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   45
            ScaleHeight     =   255
            ScaleWidth      =   1335
            TabIndex        =   46
            Top             =   540
            Width           =   1365
            Begin VB.Label Label18 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "打开(&O)"
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   45
               TabIndex        =   47
               Top             =   45
               Width           =   1680
            End
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   45
            ScaleHeight     =   255
            ScaleWidth      =   1335
            TabIndex        =   44
            Top             =   810
            Width           =   1365
            Begin VB.Label Label13 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "保存(&S)"
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   45
               TabIndex        =   45
               Top             =   45
               Width           =   1680
            End
         End
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   45
            ScaleHeight     =   255
            ScaleWidth      =   1335
            TabIndex        =   42
            Top             =   1485
            Width           =   1365
            Begin VB.Label Label14 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "退出(&C)"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   45
               TabIndex        =   43
               Top             =   45
               Width           =   1635
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   45
            ScaleHeight     =   255
            ScaleWidth      =   1335
            TabIndex        =   40
            Top             =   1125
            Width           =   1365
            Begin VB.Label Label8 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "汇入列表(&G)"
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   45
               TabIndex        =   41
               Top             =   45
               Width           =   1680
            End
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   45
            ScaleHeight     =   255
            ScaleWidth      =   1335
            TabIndex        =   38
            Top             =   270
            Width           =   1365
            Begin VB.Label Label9 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "新建(&N)"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   45
               TabIndex        =   39
               Top             =   45
               Width           =   1635
            End
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            Caption         =   "文件(&F)"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   90
            TabIndex        =   48
            Top             =   45
            Width           =   1680
         End
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   3660
         Left            =   0
         TabIndex        =   36
         Top             =   630
         Width           =   1905
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   2010
         Left            =   1935
         OLEDragMode     =   1  'Automatic
         Pattern         =   $"Cd Player1.frx":F97F
         System          =   -1  'True
         TabIndex        =   35
         ToolTipText     =   "Click to select a file"
         Top             =   630
         Width           =   2535
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   2730
         ItemData        =   "Cd Player1.frx":FA43
         Left            =   1935
         List            =   "Cd Player1.frx":FA45
         OLEDropMode     =   1  'Manual
         TabIndex        =   22
         Top             =   2700
         Width           =   2535
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   4470
         TabIndex        =   33
         Top             =   270
         Width           =   4470
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            Caption         =   "所有文件"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   3375
            TabIndex        =   70
            Top             =   90
            Width           =   1230
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            Caption         =   "媒体文件"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2250
            TabIndex        =   69
            Top             =   90
            Value           =   -1  'True
            Width           =   1230
         End
         Begin VB.DriveListBox Drive1 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   45
            TabIndex        =   34
            Top             =   45
            Width           =   1905
         End
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         X1              =   45
         X2              =   1890
         Y1              =   4545
         Y2              =   4545
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Caption         =   "选中文件"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   945
         TabIndex        =   63
         Top             =   5085
         Width           =   960
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Caption         =   "候选文件"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   945
         TabIndex        =   62
         Top             =   4815
         Width           =   960
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Caption         =   "文件夹"
         ForeColor       =   &H00FF0000&
         Height          =   600
         Left            =   45
         TabIndex        =   61
         Top             =   4815
         Width           =   870
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Caption         =   "驱动器   |文件类型"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   45
         TabIndex        =   60
         Top             =   4590
         Width           =   1860
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Caption         =   "菜单 + 工具"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   45
         TabIndex        =   59
         Top             =   4320
         Width           =   1860
      End
   End
   Begin VB.Frame Frame10 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame10"
      ForeColor       =   &H80000008&
      Height          =   5460
      Left            =   6750
      TabIndex        =   64
      Top             =   450
      Width           =   4515
   End
   Begin VB.Frame Frame12 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame12"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   90
      OLEDropMode     =   1  'Manual
      TabIndex        =   71
      Top             =   90
      Width           =   6000
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      Caption         =   "file"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1845
      TabIndex        =   68
      Top             =   8325
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label27 
      Appearance      =   0  'Flat
      Caption         =   "Label27"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2295
      TabIndex        =   67
      Top             =   8010
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      Caption         =   "unlock"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2295
      TabIndex        =   66
      Top             =   7695
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      TabIndex        =   16
      Top             =   8325
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   8055
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   7785
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Menu MumA 
      Caption         =   "A"
      Begin VB.Menu Open 
         Caption         =   "打开(&O)..."
      End
      Begin VB.Menu Add 
         Caption         =   "添加曲目(&I)..."
      End
      Begin VB.Menu Jge 
         Caption         =   "-"
      End
      Begin VB.Menu Paly 
         Caption         =   "播放(&P)"
      End
      Begin VB.Menu Pause 
         Caption         =   "暂停(&L)"
      End
      Begin VB.Menu Stop 
         Caption         =   "停止(&S)"
      End
      Begin VB.Menu Back 
         Caption         =   "上一首曲目(&B)"
      End
      Begin VB.Menu Next 
         Caption         =   "下一首曲目(&N)"
      End
      Begin VB.Menu Vol 
         Caption         =   "音量(&V)"
         Begin VB.Menu UP 
            Caption         =   "增加(&U)"
            Shortcut        =   ^U
         End
         Begin VB.Menu Do 
            Caption         =   "减少(&D)"
            Shortcut        =   ^D
         End
      End
      Begin VB.Menu jgr 
         Caption         =   "-"
      End
      Begin VB.Menu sadf 
         Caption         =   "媒体书签(&K)"
         Begin VB.Menu hg 
            Caption         =   "断点续播"
            Shortcut        =   +{INSERT}
         End
         Begin VB.Menu jgsd 
            Caption         =   "-"
         End
         Begin VB.Menu fsdf 
            Caption         =   "媒体书签 [A]"
            Shortcut        =   +^{F1}
         End
         Begin VB.Menu sdfdsf 
            Caption         =   "媒体书签 [B]"
            Shortcut        =   +^{F2}
         End
         Begin VB.Menu dsfasdf 
            Caption         =   "媒体书签 [C]"
            Shortcut        =   +^{F3}
         End
         Begin VB.Menu asdfasdf 
            Caption         =   "媒体书签 [D]"
            Shortcut        =   +^{F4}
         End
         Begin VB.Menu sdfsd 
            Caption         =   "-"
         End
         Begin VB.Menu dsfdsfds 
            Caption         =   "标记书签"
            Shortcut        =   ^{INSERT}
         End
      End
      Begin VB.Menu dsf 
         Caption         =   "-"
      End
      Begin VB.Menu Mx 
         Caption         =   "媒体类型(&T)"
         Begin VB.Menu Media 
            Caption         =   "媒体播放"
            Shortcut        =   ^{F5}
         End
         Begin VB.Menu CD 
            Caption         =   "CD 播放"
            Shortcut        =   ^{F6}
         End
      End
      Begin VB.Menu Jgx 
         Caption         =   "重复方式(&M)"
         Begin VB.Menu Lockdsf 
            Caption         =   "锁定播放"
            Shortcut        =   ^{F7}
         End
         Begin VB.Menu jghh 
            Caption         =   "-"
         End
         Begin VB.Menu Cazi 
            Caption         =   "无序播放"
            Shortcut        =   ^{F8}
         End
         Begin VB.Menu Round 
            Caption         =   "循环播放"
            Shortcut        =   ^{F9}
         End
      End
      Begin VB.Menu Jga 
         Caption         =   "视图(&W)"
         Begin VB.Menu Window 
            Caption         =   "打开视频窗口"
            Shortcut        =   +{F1}
         End
         Begin VB.Menu Full 
            Caption         =   "全屏欣赏"
            Shortcut        =   +{F2}
         End
         Begin VB.Menu JgB 
            Caption         =   "-"
         End
         Begin VB.Menu More 
            Caption         =   "完整视图"
            Shortcut        =   +{F3}
         End
         Begin VB.Menu Snow 
            Caption         =   "Snowflake"
            Shortcut        =   +{F4}
         End
         Begin VB.Menu jgf 
            Caption         =   "-"
         End
         Begin VB.Menu asdfd 
            Caption         =   "个性化播放"
            Shortcut        =   +{F5}
         End
         Begin VB.Menu dffdsaf 
            Caption         =   "-"
         End
         Begin VB.Menu OnTop 
            Caption         =   "总在最前面"
            Shortcut        =   +{F6}
         End
      End
      Begin VB.Menu Sz 
         Caption         =   "设置(&E)"
         Begin VB.Menu Setting 
            Caption         =   "播放选项..."
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu Xh 
            Caption         =   "综合设置..."
            Shortcut        =   ^{F4}
         End
         Begin VB.Menu asdfgh 
            Caption         =   "音频混合器..."
            Shortcut        =   ^{F11}
         End
      End
      Begin VB.Menu Ab 
         Caption         =   "帮助(&H)"
         Begin VB.Menu Hp 
            Caption         =   "帮助..."
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu Abb 
            Caption         =   "关于..."
            Shortcut        =   ^{F2}
         End
      End
      Begin VB.Menu jgh 
         Caption         =   "-"
      End
      Begin VB.Menu Et 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu Bdsf 
      Caption         =   "B"
      Begin VB.Menu df 
         Caption         =   "删除曲目"
         Shortcut        =   +{DEL}
      End
   End
End
Attribute VB_Name = "Form102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim T As String
Public MoveX As Integer, MoveY As Integer
Const LB_ITEMFROMPOINT = &H1A9
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public RMStop As Boolean
Dim MdNo As Integer
Dim RsS As String
Dim RmGn As Boolean, MdBo As Boolean
Dim CdOc As Boolean
Private Declare Function RasHangUp Lib "RasApi32.DLL" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
Private Declare Function RasEnumConnections Lib "RasApi32.DLL" Alias "RasEnumConnectionsA" (lprasconn As Any, lpcb As Long, lpcConnections As Long) As Long
Dim AloT As Integer
Dim SearchFlag As Integer, Cdno As Integer
Public f200 As Integer
Public jd As Integer
Public jn As String
Dim rgt As String
Public pid As Integer
Dim test As String
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
Dim SelectFileName As String
Const RAS95_MaxEntryName = 256
Const RAS95_MaxDeviceName = 128
Const RAS_MaxDeviceType = 16
Private Type RASCONN95
dwSize As Long
hRasConn As Long
szEntryName(RAS95_MaxEntryName) As Byte
szDeviceType(RAS_MaxDeviceType) As Byte
szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias _
        "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetSpecialFolderLocation Lib _
        "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder _
        As Long, pIdl As ITEMIDLIST) As Long
Private Declare Function SHGetFileInfo Lib "Shell32" Alias _
        "SHGetFileInfoA" (ByVal pszPath As Any, ByVal _
        dwFileAttributes As Long, psfi As SHFILEINFO, ByVal _
        cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
        Alias "SHGetPathFromIDListA" (ByVal pIdl As Long, ByVal _
        pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Const MAX_PATH = 260
Private Type SHITEMID
    cb As Long
    abID() As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Const SPI_SETSCREENSAVEACTIVE = 17
Const SPI_SETSCREENSAVETIMEOUT = 15
Const SPIF_SENDWININICHANGE = &H2
Const SPIF_UPDATEINIFILE = &H1

Private Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
     ByVal lpvParam As Long, ByVal fuWinIni As Long) As Long

'Private Declare Function GetVolumeInformation Lib "kernel32" _
'Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
'ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
'lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
'lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, _
'ByVal nFileSystemNameSize As Long) As Long




Sub ReMem()
On Error Resume Next
If Label26.Caption = "unlock" Then Lockdsf.Checked = False
If Label26.Caption = "locked" Then Lockdsf.Checked = True
If Check2.Value = 1 Then Cazi.Checked = True
If Check2.Value = 0 Then Cazi.Checked = False
If Check1.Value = 1 Then Round.Checked = True
If Check1.Value = 0 Then Round.Checked = False
If Check1.Value = 1 Or Check2.Value = 1 Then
Label26.Caption = "unlock"
Lockdsf.Checked = False
End If
End Sub

Private Sub SetScreenSaveTimeout(ByVal BySecond As Long)
  Call SystemParametersInfo(SPI_SETSCREENSAVETIMEOUT, BySecond, 0, _
                         SPIF_SENDWININICHANGE)
End Sub

Private Sub EnPb()
  Call SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, 1, 0, _
                        SPIF_SENDWININICHANGE)
End Sub
Private Sub DiPb()
  Call SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, 0, 0, _
                        SPIF_SENDWININICHANGE)
End Sub

Sub SetAlo()
On Error Resume Next
If Len(MediaPlayer1.FileName) > 0 Then
Form102.LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name", MediaPlayer1.FileName
Form102.LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute", MediaPlayer1.CurrentPosition
Exit Sub
End If
If Len(RA.Source) > 0 Then
Form102.LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name", Me.sReplace(RA.Source, "file://", "")
Form102.LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute", RA.GetPosition
End If
End Sub

Function sReplace(SearchLine As String, SearchFor As String, ReplaceWith As String)
On Error Resume Next

Dim vSearchLine As String, found As Integer

found = InStr(SearchLine, SearchFor): vSearchLine = SearchLine
If found <> 0 Then
vSearchLine = ""
If found > 1 Then vSearchLine = Left(SearchLine, found - 1)
vSearchLine = vSearchLine + ReplaceWith
If found + Len(SearchFor) - 1 < Len(SearchLine) Then _
vSearchLine = vSearchLine + Right$(SearchLine, Len(SearchLine) - found - Len(SearchFor) + 1)
End If
sReplace = vSearchLine

End Function

Private Function GetFolderValue(wIdx As Integer) As Long
    If wIdx < 2 Then
        GetFolderValue = 0
    ElseIf wIdx < 12 Then
        GetFolderValue = wIdx
    Else
        GetFolderValue = wIdx + 4
    End If
End Function
Private Function SendMCIString(Cmd As String, fShowError As Boolean) As Boolean
Static rc As Long
Static errStr As String * 400
rc = mciSendString(Cmd, 0, 0, hwnd)
SendMCIString = (rc = 0)
End Function
Function FileExists(FileName As String) As Boolean
On Error Resume Next
FileExists = Dir$(FileName) <> ""
If Err.Number <> 0 Then
FileExists = False
End If
On Error GoTo 0
End Function


Private Sub Abb_Click()
Call ctListBar1_ItemClick(7, 3)
End Sub


Private Sub Add_Click()
Call ctListBar1_ItemClick(2, 1)

End Sub

Private Sub asdfasdf_Click()
Call ctListBar1_ItemClick(5, 5)

End Sub

Private Sub asdfd_Click()
Call ctListBar1_ItemClick(6, 2)
End Sub

Private Sub asdfgh_Click()
Call ctListBar1_ItemClick(6, 7)
End Sub

Private Sub Back_Click()
Image5_Click
End Sub

Private Sub Cazi_Click()
Check2.Value = 1
Call ReMem
End Sub


Private Sub CD_Click()
Call Image17_Click
End Sub

Private Sub Check1_Click()

If Check1.Value = 1 Then Check2.Value = 0
Call ReMem
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then Check1.Value = 0
Call ReMem
End Sub































































Private Sub CooLine1_GotFocus()
Me.MediaPlayer1.SetFocus
End Sub


Private Sub df_Click()
If ListFile.ListCount > 0 And ListFile.SelCount > 0 Then
If pid <= ListFile.ListIndex Then pid = pid - 1
ListFile.RemoveItem (ListFile.ListIndex)
End If
End Sub

Private Sub Do_Click()
On Error Resume Next
Dim lVol As Long
Volume.Value = Volume.Value - 2000
lVol = CLng(Volume.Value) * 2
Call fSetVolumeControl(hmixer, volCtrl, lVol)
End Sub

Private Sub dsfasdf_Click()
Call ctListBar1_ItemClick(5, 4)

End Sub

Private Sub dsfdsfds_Click()
Call ctListBar1_ItemClick(1, 5)

End Sub

Private Sub Et_Click()
Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu Me.MumA, 0, x, y
End Sub



Private Sub Form_Resize()
On Error Resume Next
      If Len(SkinForm1.SkinPath) <= 0 Then GoTo cc:
If Me.WindowState = 0 Then
If Me.SkinForm1.SkinPath = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path") Then
      If myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "h", "") < 6000 Then GoTo cc
      If myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "h", "") > 7000 Then GoTo cc
      Me.Height = myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "h", "")
      Me.Width = myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "w", "")
End If
If Me.SkinForm1.SkinPath = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path") Then
      Me.Height = myReadINI(SkinForm1.SkinPath + "\sflake_info.sfl", "FORM", "h", "")
      Me.Width = myReadINI(SkinForm1.SkinPath + "\sflake_info.sfl", "FORM", "w", "")
End If
End If
'Label47.Caption = Str(Me.Height) + "   " + Str(Me.Width)
  Exit Sub
cc:
  MsgBox ("Skin 配置文件出错,请尝试重新安装 Skin 后重启 Snowman Media ilxz 3.5")
'''''End
End Sub



Private Sub Frame12_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu Me.MumA, 0, x + Frame12.Left, y + Frame12.Top

End Sub

Private Sub Frame12_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim ThisFile As Variant
    For Each ThisFile In Data.Files
        ListFile.AddItem ThisFile
    Next
End Sub

Private Sub Frame4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu Me.MumA, 0, x + Frame4.Left, y + Frame4.Top

End Sub



Private Sub Frame4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim ThisFile As Variant
    For Each ThisFile In Data.Files
        ListFile.AddItem ThisFile
    Next
End Sub

Private Sub Frame5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu Me.MumA, 0, x + Frame5.Left, y + Frame5.Top

End Sub



Private Sub Frame5_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim ThisFile As Variant
    For Each ThisFile In Data.Files
        ListFile.AddItem ThisFile
    Next
End Sub

Private Sub Frame7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu Me.MumA, 0, x + Frame7.Left, y + Frame7.Top

End Sub


Private Sub Frame7_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim ThisFile As Variant
    For Each ThisFile In Data.Files
        ListFile.AddItem ThisFile
    Next
End Sub

Private Sub fsdf_Click()
Call ctListBar1_ItemClick(5, 2)

End Sub

Private Sub Full_Click()
If f200 = 1 Then
  Form200.WindowState = 2
  Exit Sub
End If
Call ctListBar1_ItemClick(1, 3)

End Sub

Private Sub hg_Click()
Call ctListBar1_ItemClick(5, 1)
End Sub

Private Sub Hp_Click()
  If Me.FileExists(Label3.Caption + "\SmM_Hp.exe") = True Then
         Shell Label3.Caption + "\SmM_Hp.exe", vbNormalFocus
       Else
         MsgBox ("找不到文件[ " + Label3.Caption + "\SmM_Hp.exe" + " ].该文件可能已经丢失或被移动,请重新安装 Snowman Media ilxz 3.5")
       End If
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu Me.MumA, 0, x + Frame2.Left, y + Frame2.Top
If Button = 1 Then
MoveX = x
MoveY = y
End If
End Sub

Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub
Form102.Left = Form102.Left + (x - MoveX)
Form102.Top = Form102.Top + (y - MoveY)
End Sub

Private Sub Image10_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim ThisFile As Variant
    For Each ThisFile In Data.Files
        ListFile.AddItem ThisFile
    Next
                        pid = ListFile.ListCount - 1
              MediaPlayer1.FileName = ListFile.List(pid)

End Sub

Private Sub Label10_Click()
List1.Clear
End Sub
Private Sub Label11_Click()
If Len(List1.List(List1.ListIndex)) > 0 Then
If List1.ListCount > 0 And List1.SelCount > 0 Then
List1.RemoveItem (List1.ListIndex)
End If
Label28.Caption = "list"
Else: MsgBox ("还没选定要删除的曲目,请先选定好要删除的曲目再从列表中删除.")
End If
End Sub
Private Sub Label12_Click()
If Len(File1.FileName) > 0 Then
     If Len(Dir1.Path) = 3 Then
     List1.AddItem Dir1.Path & File1.FileName
     Else: List1.AddItem Dir1.Path & "\" & File1.FileName
      End If
     Else: MsgBox ("还没选定要添加的曲目,请先选定好要添加的曲目再添加入列表.")
    
End If
Label28.Caption = "file"
End Sub
Private Sub Label13_Click()
If List1.ListCount > 0 Then
  CommonDialog1.Filter = "列表文件:M3u" & _
          "|*.m3u|所有文件:*.*|*.*"
CommonDialog1.FileName = ""
CommonDialog1.ShowSave
If Len(CommonDialog1.FileName) > 0 Then
Dim i As Integer
                  If Right(CommonDialog1.FileName, 4) <> ".m3u" Then CommonDialog1.FileName = CommonDialog1.FileName + ".m3u"
   Open CommonDialog1.FileName For Output As #1
    For i = 0 To List1.ListCount - 1
     Print #1, List1.List(i)
    Next i
   Close (1)
End If
  Else: MsgBox ("当前播放列表为空无法保存,请先加入曲目到列表.")
End If
End Sub
Private Sub Label14_Click()
     If List1.ListCount > 0 Then
  CommonDialog1.Filter = "列表文件:M3u" & _
          "|*.m3u|所有文件:*.*|*.*"
CommonDialog1.FileName = ""
CommonDialog1.ShowSave
If Len(CommonDialog1.FileName) > 0 Then
Dim i As Integer
                  If Right(CommonDialog1.FileName, 4) <> ".m3u" Then CommonDialog1.FileName = CommonDialog1.FileName + ".m3u"
   Open CommonDialog1.FileName For Output As #1
    For i = 0 To List1.ListCount - 1
     Print #1, List1.List(i)
    Next i
   Close (1)
End If
  End If
         Call ToolA
End Sub
Private Sub Label18_Click()
If List1.ListCount > 0 Then
MsgBox ("当前列表中有曲目,选择保存当前列表或以取消跳过后编辑器将确认打开列表.跳过后当前所编辑列表将无法恢复.")
  CommonDialog1.Filter = "列表文件:M3u" & _
          "|*.m3u|所有文件:*.*|*.*"
CommonDialog1.FileName = ""
CommonDialog1.ShowSave
If Len(CommonDialog1.FileName) > 0 Then
Dim i As Integer
                  If Right(CommonDialog1.FileName, 4) <> ".m3u" Then CommonDialog1.FileName = CommonDialog1.FileName + ".m3u"
   Open CommonDialog1.FileName For Output As #1
    For i = 0 To List1.ListCount - 1
     Print #1, List1.List(i)
    Next i
   Close (1)
End If
End If
  CommonDialog1.Filter = "列表文件:M3u" & _
          "|*.m3u|所有文件:*.*|*.*"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
If Len(CommonDialog1.FileName) > 0 Then
List1.Clear
'Dim test As String
 Open CommonDialog1.FileName For Input As #1
    While Not EOF(1)
    Line Input #1, test
    List1.AddItem RTrim(test)
    Wend
    Close #1
End If
End Sub





Private Sub Label4_Click()
Dim i As Integer
If Len(Dir1.Path) = 3 Then
For i = 0 To File1.ListCount - 1
  List1.AddItem Dir1.Path + File1.List(i)
  Next i
Else
For i = 0 To File1.ListCount - 1
  List1.AddItem Dir1.Path + "\" + File1.List(i)
  Next i
End If
End Sub











Private Sub Label8_Click()
   CommonDialog1.FileName = ""
  CommonDialog1.Filter = "列表文件:M3u" & _
          "|*.m3u|所有文件:*.*|*.*"
             CommonDialog1.ShowOpen
             If Len(CommonDialog1.FileName) > 0 Then
              Open CommonDialog1.FileName For Input As #1
           While Not EOF(1)
          Line Input #1, test
           List1.AddItem RTrim(test)
           Wend
             Close #1
            End If
End Sub
Private Sub Label9_Click()
If List1.ListCount > 0 Then
MsgBox ("当前列表中有曲目,选择保存当前列表或以取消跳过后编辑器将自动建立新列表.跳过后当前所编辑列表将无法恢复.")
  CommonDialog1.Filter = "列表文件:M3u" & _
          "|*.m3u|所有文件:*.*|*.*"
CommonDialog1.FileName = ""
CommonDialog1.ShowSave
If Len(CommonDialog1.FileName) > 0 Then
Dim i As Integer
                  If Right(CommonDialog1.FileName, 4) <> ".m3u" Then CommonDialog1.FileName = CommonDialog1.FileName + ".m3u"
   Open CommonDialog1.FileName For Output As #1
    For i = 0 To List1.ListCount - 1
     Print #1, List1.List(i)
    Next i
   Close (1)
End If
End If
List1.Clear
End Sub



Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim ThisFile As Variant
    For Each ThisFile In Data.Files
        List1.AddItem ThisFile
    Next
End Sub




Private Sub ListFile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim pos As Long, idx As Long
pos = x / Screen.TwipsPerPixelX + y / Screen.TwipsPerPixelY * 65536
idx = SendMessage(ListFile.hwnd, LB_ITEMFROMPOINT, 0, ByVal pos)
' idx 即等于鼠标所在位置的选项
If idx < 65536 Then
ListFile.ListIndex = idx
ListFile.ToolTipText = "[" + Str(idx + 1) + " -" + Str(ListFile.ListCount) + " ] " + ListFile.List(idx)
End If
End Sub


Private Sub Lockdsf_Click()
Call ctListBar1_ItemClick(1, 4)
Call ReMem
End Sub

Private Sub Media_Click()
Call Image16_Click
End Sub


Private Sub MediaPlayer1_Error()
On Error Resume Next


'If MediaPlayer1.ErrorCode = -2147220891 Then
T = UCase(Right(ListFile.List(pid), 3))
If T = ".RA" Or T = ".RM" Or T = "RAM" Or T = ".RT" Or T = ".RP" Or T = "SMI" Or T = "MIL" Or T = "MPV" Or T = "RMM" Or T = "RTX" Then
        RMStop = False
        
        If RA.Source = "file://" + ListFile.List(pid) Or RA.Source = ListFile.List(pid) Then
            RA.DoPlay
            Exit Sub
        Else
           RA.Source = ListFile.List(pid)
        End If
    
     End If
'End If
 'RMStop = True

End Sub
Private Sub MediaPlayer1_MouseDown(Button As Integer, ShiftState As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu Me.MumA, 0, x + Frame2.Left, y + Frame2.Top
If Button = 2 Then PopupMenu Me.MumA, 0, x + Frame2.Left, y + Frame2.Top
If Button = 1 Then
MoveX = x
MoveY = y
End If
End Sub
Private Sub MediaPlayer1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub
Form102.Left = Form102.Left + (x - MoveX)
Form102.Top = Form102.Top + (y - MoveY)
End Sub
Private Sub More_Click()
Call SkinForm1_OnSkinNotify("all", "all")
End Sub

Private Sub Next_Click()
Call Image4_Click
End Sub

Private Sub OnTop_Click()
On Error Resume Next
If f200 = 1 Then
If OnTop.Checked = True Then
Form102.LyfTools1.MakeTop Form200, False
OnTop.Checked = False
Else
Form102.LyfTools1.MakeTop Form200, True
OnTop.Checked = True
   
End If
Exit Sub
End If
If OnTop.Checked = True Then
Me.LyfTools1.MakeTop Me, False
OnTop.Checked = False
Else
Me.LyfTools1.MakeTop Me, True
OnTop.Checked = True
   
End If
End Sub

Private Sub Open_Click()
Call ctListBar1_ItemClick(3, 1)
End Sub

Private Sub Option1_Click()
File1.Pattern = "*.au;*.and;*.aif;*.wmv;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.wma;*.wav;*.avi;*.mid;*.smi;*.smil;*.rt;*.mpv;*.rp;*.ram;*.rmm;*.rtx;*.ra;*.rm"
End Sub
Private Sub option1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call reco
End Sub
Private Sub Option2_Click()
File1.Pattern = "*.*"
End Sub
Private Sub option2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call reco
End Sub
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub
Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call reco
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1
End Sub
Private Sub File1_Click()
Label28.Caption = "file"
End Sub
Private Sub File1_DblClick()
If File1.ListCount > 0 Then
If Len(Dir1.Path) = 3 Then
List1.AddItem Dir1.Path & File1.FileName
Else: List1.AddItem Dir1.Path & "\" & File1.FileName
End If
End If
Label28.Caption = "file"
End Sub
Private Sub File1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 2 Then Exit Sub
        If Len(File1.FileName) > 0 Then
              If Len(Dir1.Path) = 3 Then
                    ListFile.AddItem Dir1.Path & File1.FileName
                    pid = ListFile.ListCount - 1
                    MediaPlayer1.FileName = ListFile.List(pid)
                    Else:
                    ListFile.AddItem Dir1.Path & "\" & File1.FileName
                    pid = ListFile.ListCount - 1
                    MediaPlayer1.FileName = ListFile.List(pid)
              End If
        End If
Label28.Caption = "file"
End Sub
Private Sub File1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call reco
Label28.Caption = "file"
End Sub

Sub ToolB()
      Frame2.Left = 10000
            Frame5.Left = 10000
             Frame6.Left = 10000
             Frame7.Left = 10000
              Frame9.Left = 1530
               Frame10.Left = 1575
                  List1.Clear
End Sub
Sub ToolA()
    If Frame2.Left <> 1530 Then
       Frame2.Left = 1530
            Frame5.Left = 1575
             Frame6.Left = 1530
             Frame7.Left = 1575
              Frame9.Left = 10000
               Frame10.Left = 10000
                  List1.Clear
 End If
End Sub


'Private Sub WB_DocumentComplete(ByVal pDisp As Object, url As Variant)
'If WB.Left = 0 Then CooLine1.Display = "[ H2ont Media Channer ] " + url + "  - Snowman Media ilxz 3.5"
'End Sub
Private Sub Form_Load()
On Error Resume Next
Me.MousePointer = 11
LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Starting", True
     Label3.Caption = App.Path
   If Me.FileExists(Label3.Caption + "\SmM_PT\SmM_LG.gif") = False Then
            MsgBox ("找不到文件[ " + Label3.Caption + "\SmM_PT\SmM_LG.gif" + " ].该文件可能已经丢失或被移动,请重新安装 Snowman Media ilxz 3.5")

   End
   End If
          Image10.Picture = LoadPicture(Label3.Caption + "\SmM_PT\smM_LG.gif")

If Me.FileExists(Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path") + "/skin_info.skin") = False Then
      MsgBox ("Skin 文件配置出错无法启动,请重新安装 Sm.M..")
      End
      End If
       If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "SFyn") = True Then
          Call ctListBar1_ItemClick(1, 1)
              Snow.Checked = True
        More.Checked = False
    Else
        SkinForm1.SkinPath = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path")

      End If
    Me.Show

  Open Label3.Caption + "\SmM30.dat" For Input As #1
    While Not EOF(1)
    Line Input #1, test
    ListFile.AddItem RTrim(test)
    Wend
    Close #1
      If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Op_" + Str(2)) = True Then Call Tiday
                       Dim i As Integer
                                  If Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Filename") <> "NoFile" Then
                                    pid = -1
                                      For i = 0 To ListFile.ListCount - 1
                                       If ListFile.List(i) = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Filename") Then
                                           pid = i
                                                 ListFile.ListIndex = pid
            
                                          End If
                                                   Next
                                                If pid = -1 Then
                           ListFile.AddItem Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Filename")
                            pid = ListFile.ListCount - 1
                            ListFile.ListIndex = pid
                                                End If
          MediaPlayer1.FileName = ListFile.List(pid)
         Label1.Caption = "media"
         TrackSelection.Left = 10000
         Frame8.Left = 2880
         Frame1.Left = 10000
      Form102.LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Filename", "NoFile"

            Else
          If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(14)) = 1 And Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name") <> "Error" Then Call ctListBar1_ItemClick(5, 1)
          End If
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(0)) = 1 Then Form102.LyfTools1.PlayWav Label3.Caption + "\SmM_MC\SmM_ST.wav", False

 CdOc = False
 MdBo = False
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(12)) = 1 Then MediaPlayer1.ClickToPlay = True

If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(7)) = 1 Then Call DiPb
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(8)) = 1 Then
     Me.LyfTools1.MakeTop Me, True
     OnTop.Checked = True
End If
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(11)) = 0 Then ctListBar1.SmoothScroll = False
ctListBar1.BackImage = LoadPicture(Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Te_" + Str(2)))
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Op_" + Str(15)) = True Then
CooLine1.DisplayStyle = 4
CooLine1.InsCharacter = " "
End If
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Op_" + Str(16)) = True Then
CooLine1.DisplayStyle = 2
CooLine1.InsCharacter = " "
End If
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Op_" + Str(17)) = True Then
CooLine1.DisplayStyle = 4
CooLine1.InsCharacter = ""
End If
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(1)) = 0 Then ListFile.OLEDropMode = 0
Label1.Caption = "media"
Frame1.Left = 10000
TrackSelection.Left = 10000
 Frame8.Left = 2880


    ctListBar1.AddList "当前列表"
    ctListBar1.AddList "媒体播放"
    ctListBar1.AddList "在线播放"
    ctListBar1.AddList "媒体书签"
    ctListBar1.AddList "其他功能"
    ctListBar1.AddList "帮助支持"
    ctListBar1.AddListImage 1, "Snowflake", 18
    ctListBar1.AddListImage 1, "视频窗口", 13
    ctListBar1.AddListImage 1, "全屏欣赏", 20
    ctListBar1.AddListImage 1, "锁定播放", 25
    ctListBar1.AddListImage 1, "标记书签", 28
    ctListBar1.AddListImage 1, "显示字幕", 24
    ctListBar1.AddListImage 1, "文件属性", 6
    ctListBar1.AddListImage 1, "统计信息", 5
    ctListBar1.AddListImage 2, "添加曲目", 1
    ctListBar1.AddListImage 2, "添加 URL", 31
    ctListBar1.AddListImage 2, "删除曲目", 19
    ctListBar1.AddListImage 2, "添加目录", 22
    ctListBar1.AddListImage 2, "汇入列表", 26
    ctListBar1.AddListImage 2, "预览列表", 14
    ctListBar1.AddListImage 2, "智能整理", 3
    ctListBar1.AddListImage 2, "导出保存", 4
    ctListBar1.AddListImage 2, "清空列表", 2
    ctListBar1.AddListImage 3, "打开媒体", 7
    ctListBar1.AddListImage 3, "VCD 播放", 23
    ctListBar1.AddListImage 3, "媒体光盘播放", 9
    ctListBar1.AddListImage 3, "Flash 播放", 8
    ctListBar1.AddListImage 4, "在线媒体", 10
    ctListBar1.AddListImage 4, "在线 Flash", 11
    ctListBar1.AddListImage 5, "断点续播", 16
    ctListBar1.AddListImage 5, "媒体书签 [A]", 27
    ctListBar1.AddListImage 5, "媒体书签 [B]", 27
    ctListBar1.AddListImage 5, "媒体书签 [C]", 27
    ctListBar1.AddListImage 5, "媒体书签 [D]", 27
    ctListBar1.AddListImage 6, "列表编辑器", 15
    ctListBar1.AddListImage 6, "个性化播放", 30
    ctListBar1.AddListImage 6, "功能插件", 12
    ctListBar1.AddListImage 6, "自动关机", 21
    ctListBar1.AddListImage 6, "播放选项", 17
    ctListBar1.AddListImage 6, "综合设置", 29
    ctListBar1.AddListImage 6, "音频混合器", 40
    ctListBar1.AddListImage 7, "帮助主题", 32
    ctListBar1.AddListImage 7, "检查更新", 33
    ctListBar1.AddListImage 7, "关于 Sm.M.", 34
    ctListBar1.AddListImage 7, "自述文件", 38
    ctListBar1.AddListImage 7, "许可协议", 39
    ctListBar1.AddListImage 7, "流动网络", 36
    ctListBar1.AddListImage 7, "联系作者", 35
    ctListBar1.AddListImage 7, "退出 Sm.M.", 37
'nDragPic = 0
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
If (App.PrevInstance = True) Then End
FastForwardSpeed = 10
CDLoad = False

'?????????????????????????????????????????????????????????????
SendMCIString "open cdaudio alias cd wait shareable", True
'If (SendMCIString("open cdaudio alias cd wait shareable", True) = False) Then End
SendMCIString "set cd time format tmsf wait", True

 RA.SetNoLogo True
  RA.SetEnableContextMenu False
      Unload frmAbout
      Unload Form100
      Unload Form2
      Unload Form200
      Unload Form4
      Unload Form5
      Unload Formo
    
 Dim rtn As String
Dim AllDrives As String
Dim JustOneDrive As String
Dim DriveType As Integer
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
       Label27.Caption = UCase(JustOneDrive)  'return the drive letter to the user
        Exit Do
     End If
  End If
Loop Until AllDrives = "" Or DriveType = DRIVE_CDROM


          
        If Label1.Caption = "cd" Then
   CD.Checked = True
   Media.Checked = False
      Else
   CD.Checked = False
   Media.Checked = True
        Snow.Checked = False
        More.Checked = True
        End If

 
          If Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "RunTime") = "One" Then
          Form102.LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "RunTime", "More"
          Call ctListBar1_ItemClick(7, 3)
          End If
  Timer1.Enabled = True
 Me.MousePointer = 0

End Sub



Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim ThisFile As Variant
    For Each ThisFile In Data.Files
        ListFile.AddItem ThisFile
    Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
      Me.MousePointer = 11
 Timer1.Enabled = False
     Unload frmAbout

If SkinForm1.SkinPath = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path") Then
 Form102.LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "SFyn", False
 Else
  Form102.LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "SFyn", True

 End If
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(2)) = 1 Then ListFile.Clear
Call EnPb
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Op_" + Str(3)) = True Then Call Tiday
If Len(MediaPlayer1.FileName) > 0 And Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Op_" + Str(6)) = True Then Call SetAlo
Dim i As Integer
SendMCIString "stop cd wait", True
Command = "seek cd to " & Track
SendMCIString Command, True
Playing = False
Update
SendMCIString "close all", False
Open Label3.Caption + "\SmM30.dat" For Output As #1
    For i = 0 To ListFile.ListCount - 1
     Print #1, ListFile.List(i)
    Next i
   Close (1)

LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Starting", False

End
End Sub
Private Sub Image16_Click()


If ListFile.ListCount = 0 Then
MsgBox ("同步媒体列表为空,无可用的媒体用于播放.请添加媒体曲目到同步列表后再尝试启用媒体播放模式.")
Exit Sub
End If
If Len(RA.Source) > 0 And RA.GetPosition > 0 Then
Label1.Caption = "rm"
Else
Label1.Caption = "media"
End If
Frame1.Left = 10000
TrackSelection.Left = 10000
 Frame8.Left = 2880
 If Label1.Caption = "cd" Then
   CD.Checked = True
   Media.Checked = False
Else
   CD.Checked = False
   Media.Checked = True
End If
End Sub
Private Sub Image17_Click()
On Error Resume Next
If Len(TrackTime.Caption) > 0 Then
Label1.Caption = "cd"
Frame1.Left = 45
Frame8.Left = 10000
TrackSelection.Left = 2880
Else: MsgBox ("找不到 CD 唱片,无法启用 CD 播放模式.请插入 CD 唱片后重试.")
End If
If Label1.Caption = "cd" Then
   CD.Checked = True
   Media.Checked = False
Else
   CD.Checked = False
   Media.Checked = True
End If
End Sub
Private Sub image18_Click()
 
 On Error Resume Next
 If f200 = 1 Then
     If Len(Form200.MediaPlayer1.FileName) > 0 Then
                     Form200.MediaPlayer1.Play
                   Else
                     Form200.RA.DoPlayPause
            End If
Exit Sub
End If

 RMStop = False
If Label1.Caption = "rm" Then
If RA.Left = 0 Then RMStop = False
RA.DoPlayPause
End If
If Label1.Caption = "cd" Then
SendMCIString "play cd", True
Playing = True
If Len(TrackTime.Caption) > 0 And Playing = True Then
If RA.Left = 0 Then RMStop = False
If Len(MediaPlayer1.FileName) > 0 Then MediaPlayer1.FileName = "LLXX"
If RA.Left = 0 Then RMStop = False
RA.DoStop
RA.Left = 10000
SLD1.Left = 10000
End If
End If
If Label1.Caption = "media" Then
      If pid < 0 Or pid + 1 >= ListFile.ListCount Then pid = 0
      If Len(MediaPlayer1.FileName) = 0 Then MediaPlayer1.FileName = ListFile.List(pid)
MediaPlayer1.Play
End If
End Sub
Private Sub Image3_Click()
On Error Resume Next
 If f200 = 1 Then
   If Len(Form200.MediaPlayer1.FileName) > 0 Then
   Form200.MediaPlayer1.Pause
    Else
      Form200.RA.DoPlayPause
End If
   Exit Sub
End If
RMStop = False
If Label1.Caption = "rm" Then
RA.DoPlayPause
Exit Sub
End If
If Label1.Caption = "cd" Then
SendMCIString "pause cd", True
Playing = False
Update
If Len(TrackTime.Caption) > 0 And Playing = True Then
If RA.Left = 0 Then RMStop = False
If Len(MediaPlayer1.FileName) > 0 Then MediaPlayer1.FileName = "LLXX"
If RA.Left = 0 Then RMStop = False
RA.DoStop
RA.Left = 10000
SLD1.Left = 10000
End If
Exit Sub
End If
If Label1.Caption = "media" Then
On Error Resume Next
 MediaPlayer1.Pause
If f200 = 1 Then
MediaPlayer1.FileName = "LLXX"
Form200.MediaPlayer1.Pause
End If







End If
End Sub


Private Sub image11_Click()
On Error Resume Next
MdBo = False
RMStop = False
If CdOc = False Then
SendMCIString "stop cd wait", True
Command = "seek cd to " & Track
SendMCIString Command, True
Playing = False
  If Left(MediaPlayer1.FileName, 2) = Left(Label27.Caption, 2) Then
  MdBo = False
  RMStop = False
  MediaPlayer1.FileName = "LLXX"
  End If
  If Right(Left(RA.GetSource, 9), 2) = Left(Label27.Caption, 2) And RA.GetPosition > 0 Then
  MdBo = False
  RMStop = False
  RA.DoStop
  End If
  
  SendMCIString "set cd door open", True
CdOc = True
Image11.ToolTipText = "送入光驱"
Else
 SendMCIString "set cd door closed", True
 CdOc = False
 Image11.ToolTipText = "弹出光驱"
End If
Update
End Sub
Private Sub Image7_Click()
On Error Resume Next
RMStop = False
If Label1.Caption = "rm" Then
RA.SetPosition RA.GetPosition + 5000

Exit Sub
End If
If Label1.Caption = "cd" Then
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
If Len(TrackTime.Caption) > 0 And Playing = True Then
If RA.Left = 0 Then RMStop = False
If Len(MediaPlayer1.FileName) > 0 Then MediaPlayer1.FileName = "LLXX"
If RA.Left = 0 Then RMStop = False
RA.DoStop
RA.Left = 10000
SLD1.Left = 10000
End If
Exit Sub
End If
If Label1.Caption = "media" Then
MediaPlayer1.CurrentPosition = MediaPlayer1.CurrentPosition + 5




If f200 = 1 Then
MediaPlayer1.FileName = "LLXX"
Form200.MediaPlayer1.CurrentPosition = Form200.MediaPlayer1.CurrentPosition + 5
End If


End If
End Sub
Private Sub Image6_Click()
On Error Resume Next
RMStop = False
If Label1.Caption = "rm" Then
If RA.GetPosition - 5000 < 0 Then
RA.SetPosition 0

Else: RA.SetPosition RA.GetPosition - 5000
End If
Exit Sub
End If
If Label1.Caption = "cd" Then
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
If Len(TrackTime.Caption) > 0 And Playing = True Then
If RA.Left = 0 Then RMStop = False
If Len(MediaPlayer1.FileName) > 0 Then MediaPlayer1.FileName = "LLXX"
If RA.Left = 0 Then RMStop = False
RA.DoStop
RA.Left = 10000
SLD1.Left = 10000
End If
Exit Sub
End If
If Label1.Caption = "media" Then
If MediaPlayer1.CurrentPosition - 5 < 0 Then
MediaPlayer1.CurrentPosition = -1

Else: MediaPlayer1.CurrentPosition = MediaPlayer1.CurrentPosition - 5
End If
If f200 = 1 Then
MediaPlayer1.FileName = "LLXX"
If Form200.MediaPlayer1.CurrentPosition - 5 < 0 Then
Form200.MediaPlayer1.CurrentPosition = -1

Else: Form200.MediaPlayer1.CurrentPosition = Form200.MediaPlayer1.CurrentPosition - 5
End If


End If
End If
End Sub
Private Sub Image4_Click()
On Error Resume Next
RMStop = False
If RA.Left = 0 Then RMStop = False
If Label1.Caption = "cd" Then
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
If Len(TrackTime.Caption) > 0 And Playing = True Then
If RA.Left = 0 Then RMStop = False
If Len(MediaPlayer1.FileName) > 0 Then MediaPlayer1.FileName = "LLXX"
If RA.Left = 0 Then RMStop = False
RA.DoStop
RA.Left = 10000
SLD1.Left = 10000
End If
Exit Sub
End If
If Label1.Caption = "media" Or Label1.Caption = "rm" Then
      If f200 = 0 Then
              pid = pid + 1
                  If pid = ListFile.ListCount Then
                      pid = 0
                  End If
              MediaPlayer1.FileName = ListFile.List(pid)
     End If
     If f200 = 1 Then
             MediaPlayer1.FileName = "LLXX"
                 pid = pid + 1
                     If pid = ListFile.ListCount Then
                             pid = 0
                     End If
              Form200.MediaPlayer1.FileName = ListFile.List(pid)
      End If
End If
End Sub
Private Sub Image5_Click()
'On Error Resume Next
RMStop = False
If RA.Left = 0 Then RMStop = False
If Label1.Caption = "cd" Then
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
If Len(TrackTime.Caption) > 0 And Playing = True Then
If RA.Left = 0 Then RMStop = False
If Len(MediaPlayer1.FileName) > 0 Then MediaPlayer1.FileName = "LLXX"
If RA.Left = 0 Then RMStop = False
RA.DoStop
RA.Left = 10000
SLD1.Left = 10000
End If
Exit Sub
End If
If Label1.Caption = "media" Or Label1.Caption = "rm" Then
          If f200 = 0 Then
                pid = pid - 1
                      If pid < 0 Then pid = ListFile.ListCount - 1
                MediaPlayer1.FileName = ListFile.List(pid)
          End If
            If f200 = 1 Then
           MediaPlayer1.FileName = "LLXX"
           pid = pid - 1
           If pid < 0 Then pid = 0
        Form200.MediaPlayer1.FileName = ListFile.List(pid)
End If


End If
End Sub
Private Sub Image2_Click()
On Error Resume Next
MdBo = False
RMStop = False
If Label1.Caption = "rm" Then
If RA.Left = 0 Then RMStop = False
   Call SetAlo
RA.DoStop
RA.DoStop
RA.Left = 10000
SLD1.Left = 10000
Exit Sub
End If
If Label1.Caption = "cd" Then
SendMCIString "stop cd wait", True
Command = "seek cd to " & Track
SendMCIString Command, True
Playing = False
Update
End If
If Label1.Caption = "media" Then
Call SetAlo
MediaPlayer1.FileName = "LLXX"


If f200 = 1 Then Form200.MediaPlayer1.FileName = "LLXX"







End If
End Sub
Private Function fSetVolumeControl(ByVal hmixer As Long, _
    mxc As MIXERCONTROL, ByVal Volume As Long) As Boolean
Dim rc   As Long
Dim mxcd As MIXERCONTROLDETAILS
Dim Vol  As MIXERCONTROLDETAILS_UNSIGNED
With mxcd
    .Item = 0
    .dwControlID = mxc.dwControlID
    .cbStruct = Len(mxcd)
    .cbDetails = Len(Vol)
End With
hmem = GlobalAlloc(&H40, Len(Vol))
mxcd.paDetails = GlobalLock(hmem)
mxcd.cChannels = 1
Vol.dwValue = Volume
Call CopyPtrFromStruct(mxcd.paDetails, Vol, Len(Vol))
rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
Call GlobalFree(hmem)
If MMSYSERR_NOERROR = rc Then
    fSetVolumeControl = True
Else
    fSetVolumeControl = False
End If
End Function
Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture11.BackColor = &HFFFF&
Label12.BackColor = &HFFFF&
Label12.ForeColor = &HFF0000
Picture8.BackColor = &HFFFF&
Label11.BackColor = &HFFFF&
Label11.ForeColor = &HFF0000
Picture2.BackColor = &HFFFF&
Label4.BackColor = &HFFFF&
Label4.ForeColor = &HFF0000
Picture7.BackColor = &HFF0000
Label10.BackColor = &HFF0000
Label10.ForeColor = &HFFFF&
Picture4.BackColor = &HFFFF&
Label7.BackColor = &HFFFF&
Label7.ForeColor = &HFF0000
End Sub
Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture11.BackColor = &HFFFF&
Label12.BackColor = &HFFFF&
Label12.ForeColor = &HFF0000
Picture8.BackColor = &HFF0000
Label11.BackColor = &HFF0000
Label11.ForeColor = &HFFFF&
Picture2.BackColor = &HFFFF&
Label4.BackColor = &HFFFF&
Label4.ForeColor = &HFF0000
Picture7.BackColor = &HFFFF&
Label10.BackColor = &HFFFF&
Label10.ForeColor = &HFF0000
Picture4.BackColor = &HFFFF&
Label7.BackColor = &HFFFF&
Label7.ForeColor = &HFF0000
End Sub
Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture11.BackColor = &HFF0000
Label12.BackColor = &HFF0000
Label12.ForeColor = &HFFFF&
Picture8.BackColor = &HFFFF&
Label11.BackColor = &HFFFF&
Label11.ForeColor = &HFF0000
Picture2.BackColor = &HFFFF&
Label4.BackColor = &HFFFF&
Label4.ForeColor = &HFF0000
Picture7.BackColor = &HFFFF&
Label10.BackColor = &HFFFF&
Label10.ForeColor = &HFF0000
Picture4.BackColor = &HFFFF&
Label7.BackColor = &HFFFF&
Label7.ForeColor = &HFF0000
End Sub
Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture9.BackColor = &HFF0000
Label13.BackColor = &HFF0000
Label13.ForeColor = &HFFFF&
Picture6.BackColor = &HFFFF&
Label9.BackColor = &HFFFF&
Label9.ForeColor = &HFF0000
Picture15.BackColor = &HFFFF&
Label18.BackColor = &HFFFF&
Label18.ForeColor = &HFF0000
Picture13.BackColor = &HFFFF&
Label16.BackColor = &HFFFF&
Label16.ForeColor = &HFF0000
Picture5.BackColor = &HFFFF&
Label8.BackColor = &HFFFF&
Label8.ForeColor = &HFF0000
Picture10.BackColor = &HFFFF&
Label14.BackColor = &HFFFF&
Label14.ForeColor = &HFF0000
End Sub
Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture10.BackColor = &HFF0000
Label14.BackColor = &HFF0000
Label14.ForeColor = &HFFFF&
Picture6.BackColor = &HFFFF&
Label9.BackColor = &HFFFF&
Label9.ForeColor = &HFF0000
Picture15.BackColor = &HFFFF&
Label18.BackColor = &HFFFF&
Label18.ForeColor = &HFF0000
Picture13.BackColor = &HFFFF&
Label16.BackColor = &HFFFF&
Label16.ForeColor = &HFF0000
Picture5.BackColor = &HFFFF&
Label8.BackColor = &HFFFF&
Label8.ForeColor = &HFF0000
Picture9.BackColor = &HFFFF&
Label13.BackColor = &HFFFF&
Label13.ForeColor = &HFF0000
End Sub
Private Sub Label16_Click()
Picture13.Height = 1860
Picture13.BackColor = &HFFFF&
Label16.BackColor = &HFFFF&
Label16.ForeColor = &HFF0000
End Sub
Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Picture13.Height = 285 Then
Picture13.BackColor = &HFF0000
Label16.BackColor = &HFF0000
Label16.ForeColor = &HFFFF&
End If
Picture4.Height = 285
Picture3.Height = 285
Picture4.BackColor = &HFFFF&
Label7.BackColor = &HFFFF&
Label7.ForeColor = &HFF0000
Picture3.BackColor = &HFFFF&
Label6.BackColor = &HFFFF&
Label6.ForeColor = &HFF0000
Picture6.BackColor = &HFFFF&
Label9.BackColor = &HFFFF&
Label9.ForeColor = &HFF0000
Picture15.BackColor = &HFFFF&
Label18.BackColor = &HFFFF&
Label18.ForeColor = &HFF0000
Picture5.BackColor = &HFFFF&
Label8.BackColor = &HFFFF&
Label8.ForeColor = &HFF0000
Picture9.BackColor = &HFFFF&
Label13.BackColor = &HFFFF&
Label13.ForeColor = &HFF0000
Picture10.BackColor = &HFFFF&
Label14.BackColor = &HFFFF&
Label14.ForeColor = &HFF0000
End Sub
Private Sub Label19_Click()
If Label28.Caption = "file" Then
        If Len(File1.FileName) > 0 Then
              If Len(Dir1.Path) = 3 Then
                    ListFile.AddItem Dir1.Path & File1.FileName
                    pid = ListFile.ListCount - 1
                    MediaPlayer1.FileName = ListFile.List(pid)
                    Else:
                    ListFile.AddItem Dir1.Path & "\" & File1.FileName
                    pid = ListFile.ListCount - 1
                    MediaPlayer1.FileName = ListFile.List(pid)
              End If
            Else: MsgBox ("还没有选定要预览的曲目,请先选择好要预览的曲目再预览.")
        End If
End If
If Label28.Caption = "list" Then
         If Len(List1.List(List1.ListIndex)) > 0 Then
                  ListFile.AddItem List1.List(List1.ListIndex)
                pid = ListFile.ListCount - 1
                    MediaPlayer1.FileName = ListFile.List(pid)
               Else: MsgBox ("还没有选定要预览的曲目,请先选择好要预览的曲目再预览.")
          End If
End If
End Sub
Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture3.BackColor = &HFFFF&
Label6.BackColor = &HFFFF&
Label6.ForeColor = &HFF0000
Picture16.BackColor = &HFF0000
Label19.BackColor = &HFF0000
Label19.ForeColor = &HFFFF&
Picture17.BackColor = &HFFFF&
Label20.BackColor = &HFFFF&
Label20.ForeColor = &HFF0000
End Sub
Private Sub Label20_Click()
Dim i As Integer
Dim ib As Integer
If List1.ListCount - 1 > 0 Then
  For i = 0 To List1.ListCount - 1
    ListFile.AddItem (List1.List(i))
  ib = i
  Next
MediaPlayer1.FileName = ListFile.List(ListFile.ListCount - ib)
Else: MsgBox ("列表为空无法预览,请先建立好列表再预览.")
End If
End Sub
Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture3.BackColor = &HFFFF&
Label6.BackColor = &HFFFF&
Label6.ForeColor = &HFF0000
Picture16.BackColor = &HFFFF&
Label19.BackColor = &HFFFF&
Label19.ForeColor = &HFF0000
Picture17.BackColor = &HFF0000
Label20.BackColor = &HFF0000
Label20.ForeColor = &HFFFF&
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture11.BackColor = &HFFFF&
Label12.BackColor = &HFFFF&
Label12.ForeColor = &HFF0000
Picture8.BackColor = &HFFFF&
Label11.BackColor = &HFFFF&
Label11.ForeColor = &HFF0000
Picture2.BackColor = &HFF0000
Label4.BackColor = &HFF0000
Label4.ForeColor = &HFFFF&
Picture7.BackColor = &HFFFF&
Label10.BackColor = &HFFFF&
Label10.ForeColor = &HFF0000
Picture4.BackColor = &HFFFF&
Label7.BackColor = &HFFFF&
Label7.ForeColor = &HFF0000
End Sub
Private Sub Label6_Click()
Picture3.Height = 915
Picture3.BackColor = &HFFFF&
Label6.BackColor = &HFFFF&
Label6.ForeColor = &HFF0000
Picture16.BackColor = &HFFFF&
Label19.BackColor = &HFFFF&
Label19.ForeColor = &HFF0000
Picture17.BackColor = &HFFFF&
Label20.BackColor = &HFFFF&
Label20.ForeColor = &HFF0000
End Sub
Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Picture3.Height = 285 Then
Picture3.BackColor = &HFF0000
Label6.BackColor = &HFF0000
Label6.ForeColor = &HFFFF&
End If
Picture13.Height = 285
Picture4.Height = 285
Picture13.BackColor = &HFFFF&
Label16.BackColor = &HFFFF&
Label16.ForeColor = &HFF0000
Picture4.BackColor = &HFFFF&
Label7.BackColor = &HFFFF&
Label7.ForeColor = &HFF0000
End Sub
Private Sub Label7_Click()
Picture4.Height = 1500
Picture4.BackColor = &HFFFF&
Label7.BackColor = &HFFFF&
Label7.ForeColor = &HFF0000
End Sub
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Picture4.Height = 285 Then
Picture4.BackColor = &HFF0000
Label7.BackColor = &HFF0000
Label7.ForeColor = &HFFFF&
End If
Picture13.Height = 285
Picture3.Height = 285
Picture13.BackColor = &HFFFF&
Label16.BackColor = &HFFFF&
Label16.ForeColor = &HFF0000
Picture3.BackColor = &HFFFF&
Label6.BackColor = &HFFFF&
Label6.ForeColor = &HFF0000
End Sub
Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture5.BackColor = &HFF0000
Label8.BackColor = &HFF0000
Label8.ForeColor = &HFFFF&
Picture6.BackColor = &HFFFF&
Label9.BackColor = &HFFFF&
Label9.ForeColor = &HFF0000
Picture15.BackColor = &HFFFF&
Label18.BackColor = &HFFFF&
Label18.ForeColor = &HFF0000
Picture13.BackColor = &HFFFF&
Label16.BackColor = &HFFFF&
Label16.ForeColor = &HFF0000
Picture9.BackColor = &HFFFF&
Label13.BackColor = &HFFFF&
Label13.ForeColor = &HFF0000
Picture10.BackColor = &HFFFF&
Label14.BackColor = &HFFFF&
Label14.ForeColor = &HFF0000
End Sub
Private Sub List1_Click()
Label28.Caption = "list"
End Sub
Private Sub List1_DblClick()
If List1.ListCount > 0 And List1.SelCount > 0 Then
List1.RemoveItem (List1.ListIndex)
End If
Label28.Caption = "list"
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 2 Then Exit Sub
  If Len(List1.List(List1.ListIndex)) > 0 Then
                  ListFile.AddItem List1.List(List1.ListIndex)
                pid = ListFile.ListCount - 1
                    MediaPlayer1.FileName = ListFile.List(pid)
          End If
Label28.Caption = "list"
End Sub
Private Sub ListFile_DblClick()
MdBo = False
pid = ListFile.ListIndex

   'T = UCase(Right(ListFile.List(pid), 3))
     'If T = ".RA" Or T = ".RM" Or T = "RAM" Or T = ".RT" Or T = ".RP" Or T = "SMI" Or T = "MIL" Then
     '   If RA.Source = "file://" + ListFile.List(pid) Or RA.Source = ListFile.List(pid) Then
      '      RMStop = False
      '      RA.DoPlay
      '      Exit Sub
      '  Else
    '       RA.Source = ListFile.List(pid)
  '      End If
'    End If
 If Me.FileExists(ListFile.List(pid)) = False Then
   MsgBox ("找不到指定的媒体文件,该文件可能已经丢失或被移动.请确定该文件存在后重试.")
 Exit Sub
 End If
 
  If Right(ListFile.List(pid), 4) = ".cda" Then
    Call Image17_Click
    If (CDLoad) Then
        If (Track <= TotalTracks) Then
            If (Playing) Then
                Command = "play cd from " & Val(Left(Right(ListFile.List(pid), 6), 2))
                SendMCIString Command, True
             Else
                Command = "seek cd to " & Val(Left(Right(ListFile.List(pid), 6), 2))
                SendMCIString Command, True
                SendMCIString "play cd", True
                Playing = True
            End If
        End If
       Else
        SendMCIString "seek cd to 1", True
    End If
    Update
    If Len(TrackTime.Caption) > 0 And Playing = True Then
If RA.Left = 0 Then RMStop = False
If Len(MediaPlayer1.FileName) > 0 Then MediaPlayer1.FileName = "LLXX"
If RA.Left = 0 Then RMStop = False
RA.DoStop
RA.Left = 10000
SLD1.Left = 10000
End If


  Exit Sub
  End If
MediaPlayer1.FileName = ListFile.List(pid)
 If f200 = 1 Then
 MediaPlayer1.FileName = "LLXX"
 Form200.MediaPlayer1.FileName = ListFile.List(pid)
 MediaPlayer1.FileName = "LLXX"
End If
Label1.Caption = "media"
Frame1.Left = 10000
TrackSelection.Left = 10000
 Frame8.Left = 2880

End Sub
Private Sub ListFile_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu Me.Bdsf, 0, Frame6.Left + x, Frame6.Top + y
End Sub
Private Sub Listfile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim ThisFile As Variant
    For Each ThisFile In Data.Files
        ListFile.AddItem ThisFile
    Next
End Sub
Sub mpn()
If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(3)) = 1 Then
Dim lngRetCode As Long
Dim lpcb As Long
Dim lpcConnections As Long
Dim intArraySize As Integer
Dim intLooper As Integer
ReDim lprasconn95(intArraySize) As RASCONN95
lprasconn95(0).dwSize = 412
lpcb = 256 * lprasconn95(0).dwSize
lngRetCode = RasEnumConnections(lprasconn95(0), lpcb, lpcConnections)
If lngRetCode = 0 Then
If lpcConnections > 0 Then
For intLooper = 0 To lpcConnections - 1
RasHangUp lprasconn95(intLooper).hRasConn
Next intLooper
End If
End If
End If
If Label26.Caption = "locked" Then
MediaPlayer1.FileName = MediaPlayer1.FileName
Exit Sub
End If
If Check1.Value = 0 And Check2.Value = 0 Then
pid = pid + 1
If pid = ListFile.ListCount Then
MediaPlayer1.FileName = "LLXX"
     Form4.Asd = True
End If
MediaPlayer1.FileName = ListFile.List(pid)
Exit Sub
End If
If Check1.Value = 1 Then
pid = pid + 1
If pid = ListFile.ListCount Then
pid = 0
End If
MediaPlayer1.FileName = ListFile.List(pid)
Exit Sub
End If
If Check2.Value = 1 Then
 Randomize
    pid = Int(ListFile.ListCount * Rnd)
MediaPlayer1.FileName = ListFile.List(pid)
Exit Sub
End If
End Sub
Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
On Error Resume Next
     If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(3)) = 1 Then
 
     Dim lngRetCode As Long
Dim lpcb As Long
Dim lpcConnections As Long
Dim intArraySize As Integer
Dim intLooper As Integer
ReDim lprasconn95(intArraySize) As RASCONN95
lprasconn95(0).dwSize = 412
lpcb = 256 * lprasconn95(0).dwSize
lngRetCode = RasEnumConnections(lprasconn95(0), lpcb, lpcConnections)
If lngRetCode = 0 Then
If lpcConnections > 0 Then
For intLooper = 0 To lpcConnections - 1
RasHangUp lprasconn95(intLooper).hRasConn
Next intLooper
End If
End If
End If












If Label26.Caption = "locked" Then
MediaPlayer1.FileName = MediaPlayer1.FileName
Exit Sub
End If

If Check1.Value = 0 And Check2.Value = 0 Then
pid = pid + 1
If pid = ListFile.ListCount Then
MediaPlayer1.FileName = "LLXX"
  Form4.Asd = True
End If
MediaPlayer1.FileName = ListFile.List(pid)
Exit Sub
End If

If Check1.Value = 1 Then
pid = pid + 1
If pid = ListFile.ListCount Then
pid = 0
End If
MediaPlayer1.FileName = ListFile.List(pid)
Exit Sub
End If

If Check2.Value = 1 Then
 Randomize
    pid = Int(ListFile.ListCount - 1 * Rnd)
MediaPlayer1.FileName = ListFile.List(pid)
Exit Sub
End If

End Sub
Private Sub MediaPlayer1_NewStream()
On Error Resume Next
'If Right(MediaPlayer1.Filename, 4) = ".m3u" Then Call M3uTd
If pid + 1 <= ListFile.ListCount Then ListFile.ListIndex = pid
RMStop = False
RA.DoStop
RA.Left = 10000
SLD1.Left = 10000
SendMCIString "stop cd wait", True
Command = "seek cd to " & Track
SendMCIString Command, True
Playing = False
Update
End Sub
Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture6.BackColor = &HFF0000
Label9.BackColor = &HFF0000
Label9.ForeColor = &HFFFF&
Picture15.BackColor = &HFFFF&
Label18.BackColor = &HFFFF&
Label18.ForeColor = &HFF0000
Picture13.BackColor = &HFFFF&
Label16.BackColor = &HFFFF&
Label16.ForeColor = &HFF0000
Picture5.BackColor = &HFFFF&
Label8.BackColor = &HFFFF&
Label8.ForeColor = &HFF0000
Picture9.BackColor = &HFFFF&
Label13.BackColor = &HFFFF&
Label13.ForeColor = &HFF0000
Picture10.BackColor = &HFFFF&
Label14.BackColor = &HFFFF&
Label14.ForeColor = &HFF0000
End Sub
Private Sub label18_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture15.BackColor = &HFF0000
Label18.BackColor = &HFF0000
Label18.ForeColor = &HFFFF&
Picture6.BackColor = &HFFFF&
Label9.BackColor = &HFFFF&
Label9.ForeColor = &HFF0000
Picture13.BackColor = &HFFFF&
Label16.BackColor = &HFFFF&
Label16.ForeColor = &HFF0000
Picture5.BackColor = &HFFFF&
Label8.BackColor = &HFFFF&
Label8.ForeColor = &HFF0000
Picture9.BackColor = &HFFFF&
Label13.BackColor = &HFFFF&
Label13.ForeColor = &HFF0000
Picture10.BackColor = &HFFFF&
Label14.BackColor = &HFFFF&
Label14.ForeColor = &HFF0000
End Sub

Private Sub Paly_Click()
Call image18_Click
End Sub

Private Sub Pause_Click()
Call Image3_Click
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call reco
End Sub
Private Sub Picture13_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture9.BackColor = &HFFFF&
Label13.BackColor = &HFFFF&
Label13.ForeColor = &HFF0000
Picture6.BackColor = &HFFFF&
Label9.BackColor = &HFFFF&
Label9.ForeColor = &HFF0000
Picture15.BackColor = &HFFFF&
Label18.BackColor = &HFFFF&
Label18.ForeColor = &HFF0000
Picture13.BackColor = &HFFFF&
Label16.BackColor = &HFFFF&
Label16.ForeColor = &HFF0000
Picture5.BackColor = &HFFFF&
Label8.BackColor = &HFFFF&
Label8.ForeColor = &HFF0000
Picture10.BackColor = &HFFFF&
Label14.BackColor = &HFFFF&
Label14.ForeColor = &HFF0000
End Sub
Sub reco()
Picture13.Height = 285
Picture4.Height = 285
Picture3.Height = 285
Picture9.BackColor = &HFFFF&
Label13.BackColor = &HFFFF&
Label13.ForeColor = &HFF0000
Picture6.BackColor = &HFFFF&
Label9.BackColor = &HFFFF&
Label9.ForeColor = &HFF0000
Picture15.BackColor = &HFFFF&
Label18.BackColor = &HFFFF&
Label18.ForeColor = &HFF0000
Picture13.BackColor = &HFFFF&
Label16.BackColor = &HFFFF&
Label16.ForeColor = &HFF0000
Picture5.BackColor = &HFFFF&
Label8.BackColor = &HFFFF&
Label8.ForeColor = &HFF0000
Picture10.BackColor = &HFFFF&
Label14.BackColor = &HFFFF&
Label14.ForeColor = &HFF0000
Picture11.BackColor = &HFFFF&
Label12.BackColor = &HFFFF&
Label12.ForeColor = &HFF0000
Picture8.BackColor = &HFFFF&
Label11.BackColor = &HFFFF&
Label11.ForeColor = &HFF0000
Picture2.BackColor = &HFFFF&
Label4.BackColor = &HFFFF&
Label4.ForeColor = &HFF0000
Picture7.BackColor = &HFFFF&
Label10.BackColor = &HFFFF&
Label10.ForeColor = &HFF0000
Picture4.BackColor = &HFFFF&
Label7.BackColor = &HFFFF&
Label7.ForeColor = &HFF0000
Picture3.BackColor = &HFFFF&
Label6.BackColor = &HFFFF&
Label6.ForeColor = &HFF0000
Picture16.BackColor = &HFFFF&
Label19.BackColor = &HFFFF&
Label19.ForeColor = &HFF0000
Picture17.BackColor = &HFFFF&
Label20.BackColor = &HFFFF&
Label20.ForeColor = &HFF0000
End Sub
Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture3.BackColor = &HFFFF&
Label6.BackColor = &HFFFF&
Label6.ForeColor = &HFF0000
Picture16.BackColor = &HFFFF&
Label19.BackColor = &HFFFF&
Label19.ForeColor = &HFF0000
Picture17.BackColor = &HFFFF&
Label20.BackColor = &HFFFF&
Label20.ForeColor = &HFF0000
End Sub
Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture11.BackColor = &HFFFF&
Label12.BackColor = &HFFFF&
Label12.ForeColor = &HFF0000
Picture8.BackColor = &HFFFF&
Label11.BackColor = &HFFFF&
Label11.ForeColor = &HFF0000
Picture2.BackColor = &HFFFF&
Label4.BackColor = &HFFFF&
Label4.ForeColor = &HFF0000
Picture7.BackColor = &HFFFF&
Label10.BackColor = &HFFFF&
Label10.ForeColor = &HFF0000
Picture4.BackColor = &HFFFF&
Label7.BackColor = &HFFFF&
Label7.ForeColor = &HFF0000
End Sub


Public Sub RA_OnClipClosed()
RA.Left = 10000
SLD1.Left = 10000
Frame11.Left = 10000
If RMStop = True Then

     
     If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(3)) = 1 Then
 Dim lngRetCode As Long
Dim lpcb As Long
Dim lpcConnections As Long
Dim intArraySize As Integer
Dim intLooper As Integer
ReDim lprasconn95(intArraySize) As RASCONN95
lprasconn95(0).dwSize = 412
lpcb = 256 * lprasconn95(0).dwSize
lngRetCode = RasEnumConnections(lprasconn95(0), lpcb, lpcConnections)
If lngRetCode = 0 Then
If lpcConnections > 0 Then
For intLooper = 0 To lpcConnections - 1
RasHangUp lprasconn95(intLooper).hRasConn
Next intLooper
End If
End If
End If
If Label26.Caption = "locked" Then
MediaPlayer1.FileName = ListFile.List(pid)
RMStop = True
Exit Sub
End If

If Check1.Value = 0 And Check2.Value = 0 Then
pid = pid + 1
If pid = ListFile.ListCount Then
MediaPlayer1.FileName = "LLXX"
     Form4.Asd = True
End If
MediaPlayer1.FileName = ListFile.List(pid)
RMStop = True
Exit Sub
End If

If Check1.Value = 1 Then
pid = pid + 1
If pid = ListFile.ListCount Then
pid = 0
End If
MediaPlayer1.FileName = ListFile.List(pid)
RMStop = True
Exit Sub
End If

If Check2.Value = 1 Then
 Randomize
    pid = Int(ListFile.ListCount * Rnd)
MediaPlayer1.FileName = ListFile.List(pid)
RMStop = True
Exit Sub
End If
End If
RMStop = True
End Sub

Private Sub RA_OnClipOpened(ByVal shortClipName As String, ByVal url As String)
On Error Resume Next
Frame11.Left = 45

Call BigSma

If pid + 1 <= ListFile.ListCount Then ListFile.ListIndex = pid
SLD1.Max = RA.GetLength
Label1.Caption = "rm"
RA.Left = 0
SLD1.Left = -45
'MediaPlayer1.Filename = "LLXX"
SendMCIString "stop cd wait", True
Command = "seek cd to " & Track
SendMCIString Command, True
Playing = False
Update
RMStop = True
 RmGn = True
End Sub
Sub BigSma()
If SkinForm1.SkinPath = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path") Then
RA.Visible = False
If RA.GetClipWidth > 0 Then
If RA.GetClipWidth / RA.GetClipHeight >= myReadINI(SkinForm1.SkinPath + "\sflake_info.sfl", "video", "w", "") / myReadINI(SkinForm1.SkinPath + "\sflake_info.sfl", "video", "h", "") Then
RA.Width = myReadINI(SkinForm1.SkinPath + "\sflake_info.sfl", "video", "w", "")

RA.Height = myReadINI(SkinForm1.SkinPath + "\sflake_info.sfl", "video", "w", "") * RA.GetClipHeight / RA.GetClipWidth
RA.Top = (myReadINI(SkinForm1.SkinPath + "\sflake_info.sfl", "video", "h", "") + RA.Height) / 2 - RA.Height
RA.Left = (myReadINI(SkinForm1.SkinPath + "\sflake_info.sfl", "video", "w", "") + RA.Width) / 2 - RA.Width
Else
RA.Height = myReadINI(SkinForm1.SkinPath + "\sflake_info.sfl", "video", "h", "")
RA.Width = myReadINI(SkinForm1.SkinPath + "\sflake_info.sfl", "video", "w", "")
'RA.Width = myReadINI(SkinForm1.SkinPath + "\sflake_info.sfl", "video", "h", "") * RA.GetClipWidth / RA.GetClipHeight
RA.Top = (myReadINI(SkinForm1.SkinPath + "\sflake_info.sfl", "video", "h", "") + RA.Height) / 2 - RA.Height
RA.Left = (myReadINI(SkinForm1.SkinPath + "\sflake_info.sfl", "video", "w", "") + RA.Width) / 2 - RA.Width
End If
RA.Visible = True
End If
Exit Sub
End If

RA.Visible = False
If RA.GetClipWidth > 0 Then
If RA.GetClipWidth / RA.GetClipHeight >= 4500 / 3375 Then
RA.Width = 4500

RA.Height = 4500 * RA.GetClipHeight / RA.GetClipWidth
RA.Top = (3375 + RA.Height) / 2 - RA.Height
RA.Left = (4500 + RA.Width) / 2 - RA.Width
Else
RA.Height = 3375
RA.Width = 4500
'RA.Width = 3375 * RA.GetClipWidth / RA.GetClipHeight
RA.Top = (3375 + RA.Height) / 2 - RA.Height
RA.Left = (4500 + RA.Width) / 2 - RA.Width
End If
RA.Visible = True
End If
End Sub
Private Sub RA_OnPositionChange(ByVal lPos As Long, ByVal lLen As Long)
If RmGn = True Then
SLD1.Value = RA.GetPosition
Dim RTa, Rtb As Integer
RTa = Int(RA.GetPosition / 1000 / 60)
Rtb = Int((RA.GetPosition - RTa * 60 * 1000) / 1000)
Label5.Caption = "[ Real Media ]  " + Right(Str(100 + RTa), 2) + ":" + Right(Str(100 + Rtb), 2) + " / "
RTa = Int(RA.GetLength / 1000 / 60)
Rtb = Int((RA.GetLength - RTa * 60 * 1000) / 1000)
Label5.Caption = Label5.Caption + Right(Str(100 + RTa), 2) + ":" + Right(Str(100 + Rtb), 2) + " "
Label5.Caption = Label5.Caption + Str(RA.GetPosition) + " /" + Str(RA.GetLength)
Label5.Left = (Label5.Width + Frame11.Width) / 2 - Label5.Width
End If





RsS = ""
   RsS = "[" + Str(pid + 1) + " -" + Str(ListFile.ListCount) + " ] "
    If RA.GetBandwidthCurrent <> 0 Then RsS = RsS + "正在播放:" + Str(Int(RA.GetBandwidthCurrent / 1000)) + " Kbps  "
    If RA.GetClipWidth > 0 Then RsS = RsS + "视频:" + Str(RA.GetClipWidth) + " *" + Str(RA.GetClipHeight) + "  "
     If Len(RA.GetTitle) > 0 Then RsS = RsS + "标题:" + RA.GetTitle + "  "
       If Len(RA.GetCopyright) > 0 And RA.GetCopyright <> "?000" Then RsS = RsS + "版权:" + RA.GetCopyright + "  "
       RsS = RsS + "地址:" + ListFile.List(pid) + "  "
         RsS = RsS + " - Snowman Media ilxz 3.5"
 If CooLine1.Display <> RsS Then CooLine1.Display = RsS
 
End Sub





Private Sub Round_Click()
Check1.Value = 1
Call ReMem
End Sub

Private Sub sdfdsf_Click()
Call ctListBar1_ItemClick(5, 3)

End Sub

Private Sub Setting_Click()
MediaPlayer1.ShowDialog mpShowDialogOptions
End Sub

Private Sub SLD1_Click()
RmGn = True
End Sub

Private Sub SLD1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RmGn = False
End Sub

Private Sub SLD1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RA.SetPosition SLD1.Value
RmGn = True
End Sub

Private Sub Snow_Click()
Call ctListBar1_ItemClick(1, 1)
End Sub

Private Sub Stop_Click()
Image2_Click
End Sub

Private Sub TrackSelection_Click()
MdBo = False
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
    If Len(TrackTime.Caption) > 0 And Playing = True Then
If RA.Left = 0 Then RMStop = False
If Len(MediaPlayer1.FileName) > 0 Then MediaPlayer1.FileName = "LLXX"
If RA.Left = 0 Then RMStop = False
RA.DoStop
RA.Left = 10000
SLD1.Left = 10000
End If
    
    
End Sub

Private Sub UP_Click()
On Error Resume Next
Dim lVol As Long
Volume.Value = Volume.Value + 2000
lVol = CLng(Volume.Value) * 2
Call fSetVolumeControl(hmixer, volCtrl, lVol)
End Sub

Private Sub Volume_Change()
Dim lVol As Long
lVol = CLng(Volume.Value) * 2
Call fSetVolumeControl(hmixer, volCtrl, lVol)
End Sub

Private Sub Volume_GotFocus()
Dim lVol As Long
lVol = CLng(Volume.Value) * 2
Call fSetVolumeControl(hmixer, volCtrl, lVol)

End Sub

Private Sub Volume_Scroll()
Dim lVol As Long
lVol = CLng(Volume.Value) * 2
Call fSetVolumeControl(hmixer, volCtrl, lVol)
End Sub
Private Sub Update()
On Error Resume Next
Static e As String * 30
mciSendString "status cd media present", e, Len(e), 0
If (CBool(e)) Then
    If (CDLoad = False) Then
        mciSendString "status cd number of tracks wait", e, Len(e), 0
        TotalTracks = CInt(Mid$(e, 1, 2))
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
     If (CDLoad = True) Then
        CDLoad = False
        Playing = False
        TrackTime.Caption = ""
        Me.TotalTrack.Caption = ""
        TimeWindow.Text = ""
    End If
End If
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
rgt = Time
  Update
AloT = AloT + 1
If AloT > 600 Then
'If AloT > 10 Then
  If Len(MediaPlayer1.FileName) > 0 And Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Op_" + Str(5)) = True Then Call SetAlo
AloT = 0
End If
  
'If MediaPlayer1.DisplaySize <> mpFullScreen And RA.GetFullScreen <> True Then
'If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(8)) = 1 Then Me.LyfTools1.MakeTop Me, True
'End If

'If Len(MediaPlayer1.Filename) > 0 Then Label1.Caption = "media"
'If Playing = True Then Label1.Caption = "cd"
If MdBo = True Then
     If pid + 1 > ListFile.ListCount Then
  MdBo = False
  pid = 0
  ListFile.ListIndex = 0
   MediaPlayer1.FileName = "LLXX"
   RMStop = False
   
   
   RA.DoStop
    CooLine1.Display = "[ 信息 ] 同步媒体列表预览完成   - Snowman Media ilxz 3.5"

  Exit Sub
    End If
  MdNo = MdNo + 1
  If MdNo >= 10 Then
        RMStop = False
        pid = pid + 1
        MediaPlayer1.FileName = ListFile.List(pid)
        MdNo = 0
  End If

End If
If Len(MediaPlayer1.FileName) > 0 Then
RsS = ""
 RsS = "[" + Str(pid + 1) + " -" + Str(ListFile.ListCount) + " ] "
     If MediaPlayer1.Bandwidth > 0 Then RsS = RsS + "正在播放:" + Str(Int(MediaPlayer1.Bandwidth / 1000)) + " Kbps  "
     If MediaPlayer1.ImageSourceWidth > 0 Then RsS = RsS + "视频:" + Str(MediaPlayer1.ImageSourceWidth) + " *" + Str(MediaPlayer1.ImageSourceHeight) + "  "
  If Len(MediaPlayer1.GetMediaInfoString(mpClipTitle)) > 0 Then RsS = RsS + "作品:" + MediaPlayer1.GetMediaInfoString(mpClipTitle) + "  "
  If Len(MediaPlayer1.GetMediaInfoString(mpClipAuthor)) > 0 Then RsS = RsS + "艺术家:" + MediaPlayer1.GetMediaInfoString(mpClipAuthor) + "  "
    If Len(MediaPlayer1.GetMediaInfoString(mpClipCopyright)) > 0 Then RsS = RsS + "版权:" + MediaPlayer1.GetMediaInfoString(mpClipCopyright) + "  "
    If Len(MediaPlayer1.GetMediaInfoString(mpClipDescription)) > 0 Then RsS = RsS + "描述:" + MediaPlayer1.GetMediaInfoString(mpClipDescription) + "  "
   RsS = RsS + "地址:" + MediaPlayer1.FileName + "  "
   RsS = RsS + " - Snowman Media ilxz 3.5"
   If CooLine1.Display <> RsS Then CooLine1.Display = RsS
End If
If Playing = True Then
  RsS = "[ CD Audio ] 艺术家:未知艺术家  唱片集:未知唱片集  标号:" + Me.LyfTools1.GetDiskNumber(Label27.Caption) + "  驱动器:" + sReplace(Label27.Caption, ":\", "") + "   - Snowman Media ilxz 3.5"
If CooLine1.Display <> RsS Then CooLine1.Display = RsS
End If
 If Len(MediaPlayer1.FileName) = 0 And RA.GetPosition = 0 And Playing = False And Me.CooLine1.Display <> "H2ont Leask Snowman Media ilxz 3.5 Plus Edition Ready Now！   " Then Me.CooLine1.Display = "H2ont Leask Snowman Media ilxz 3.5 Plus Edition Ready Now！   "
      
      
Dim i As Integer
         If Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Filename") <> "NoFile" Then
             pid = -1
                For i = 0 To ListFile.ListCount - 1
                      If ListFile.List(i) = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Filename") Then
                         pid = i
                         ListFile.ListIndex = pid
            
                      End If
                Next
                      If pid = -1 Then
                           ListFile.AddItem Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Filename")
                            pid = ListFile.ListCount - 1
                            ListFile.ListIndex = pid
                            End If
          MediaPlayer1.FileName = ListFile.List(pid)
         Label1.Caption = "media"
         TrackSelection.Left = 10000
         Frame8.Left = 2880
         Frame1.Left = 10000
         Form102.LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Filename", "NoFile"
         


End If
LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Starting", True
End Sub

Sub Tiday()
On Error Resume Next
Me.MousePointer = 11
Dim tid, tidb As Integer
tid = 0
tidb = 0
For tid = 0 To ListFile.ListCount - 1
    If FileExists(ListFile.List(tid)) = False Then
       For tidb = tid + 1 To ListFile.ListCount - 1
          If ListFile.List(tidb) = ListFile.List(tid) Then
          ListFile.RemoveItem (tidb)
          tidb = tidb - 1
          End If
        Next tidb
        ListFile.RemoveItem (tid)
    tid = tid - 1
    End If
Next tid
For tid = 0 To ListFile.ListCount - 1
    For tidb = tid + 1 To ListFile.ListCount - 1
        If ListFile.List(tidb) = ListFile.List(tid) Then
         ListFile.RemoveItem (tidb)
        tidb = tidb - 1
        End If
      Next tidb
Next tid
Me.MousePointer = 0
End Sub
Private Function DirDiver(NewPath As String, dircount As Integer, BackUp As String) As Integer
Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String
Dim retval As Integer
  SearchFlag = True             ' Set flag so user can interrupt.
  DirDiver = False              ' Set to TRUE if there is an error.
  retval = DoEvents()           ' check for events (i.e. user Cancels).
  If SearchFlag = False Then
    DirDiver = True
    Exit Function
  End If
  On Local Error GoTo DirDriverHandler
  DirsToPeek = Dir1.ListCount            ' How many directories below this?
  Do While DirsToPeek > 0 And SearchFlag = True
    OldPath = Dir1.Path                  ' Save old path for next recursion.
    Dir1.Path = NewPath
    If Dir1.ListCount > 0 Then
    ' Get to the node bottom.
      Dir1.Path = Dir1.List(DirsToPeek - 1)
      AbandonSearch = DirDiver((Dir1.Path), dircount%, OldPath)
    End If
    ' Go up 1 level in directories.
    DirsToPeek = DirsToPeek - 1
    If AbandonSearch = True Then Exit Function
  Loop
  ' Call function to enumerate files.
  If File1.ListCount Then
    If Len(Dir1.Path) <= 3 Then
        ThePath = Dir1.Path         ' If at root level, leave as is...
    Else
        ThePath = Dir1.Path + "\"   ' otherwise put "\" before file name.
    End If
    For ind = 0 To File1.ListCount - 1        ' Add conforming files in
        entry = ThePath + File1.List(ind)     ' this directory to listbox.
        ListFile.AddItem entry
       Cdno = Cdno + 1
    Next ind
  End If
  If BackUp <> "" Then         ' If there is a superior
      Dir1.Path = BackUp    ' directory, move to it.
  End If
  Exit Function
DirDriverHandler:
  If Err = 7 Then         ' If Out of Memory, assume listbox just got full.
    DirDiver = True       ' Create Msg$ and set return value AbandonSearch.
    MsgBox "You've filled the listbox. Search being abandoned..."
    Exit Function         ' Note that EXIT procedure resets ERR to 0.
  Else                    ' Otherwise display error message and quit.
    MsgBox Error
    End
  End If
End Function
Private Sub ctListBar1_ListChange(ByVal nList As Integer)
    If (nList = 1) Then
            Call ToolA
 End If
 End Sub
Private Sub SkinForm1_OnSkinNotify(ByVal SkinClass As String, ByVal SkinEvent As String)
    Select Case SkinClass
    Case "A"
        Call image18_Click
    Case "B"
        Call Image3_Click
    Case "C"
        Call Image5_Click
    Case "D"
        Call Image6_Click
    Case "E"
         Call Image7_Click
    Case "F"
         Call Image4_Click
    Case "G"
         Call Image2_Click
    Case "H"
         Call image11_Click
    Case "I"
         Call Image16_Click
    Case "J"
         Call Image17_Click
    Case "min"
         Me.WindowState = 1
    Case "all"

          If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(8)) = 1 Then
Me.LyfTools1.MakeTop Me, True
               OnTop.Checked = True
          Else
           Me.LyfTools1.MakeTop Me, False
 
                 OnTop.Checked = False
          End If
   If Me.FileExists(Label3.Caption + "\SmM_PT\SmM_LG.gif") = False Then End
          Image10.Picture = LoadPicture(Label3.Caption + "\SmM_PT\smM_LG.gif")
          SkinForm1.SkinPath = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path")
          MediaPlayer1.ShowControls = True
          MediaPlayer1.ShowStatusBar = True
          CooLine1.Left = 45
          CooLine1.Top = 45
          CooLine1.Width = 6000
          Frame12.Left = 90
          Frame4.Left = 90
          ctListBar1.Left = 45
          Frame2.Top = 405
          Frame2.Left = 1530
          Frame2.Width = 4515
          Frame2.Height = 4470
          Frame5.Left = 1575
          Frame6.Left = 1530
          Frame7.Left = 1575
          MediaPlayer1.Width = 4515
          MediaPlayer1.Height = 4470
          Snow.Checked = False
          More.Checked = True
          Call BigSma

    End Select
   
End Sub

 Private Sub ctListBar1_ItemClick(ByVal nList As Integer, ByVal nItem As Integer)
On Error Resume Next

 Dim Result As Long
Select Case nList
        Case 1
          Select Case nItem
          Case 1
               If Me.FileExists(Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path") + "\sflake_info.sfl") = False Then
      MsgBox ("Snowflake 文件配置出错无法更改视图,请重新安装 Sm.M..")
      Exit Sub
      End If
                    If Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(9)) = 1 Then
         Me.LyfTools1.MakeTop Me, True
              OnTop.Checked = True
          Else
           Me.LyfTools1.MakeTop Me, False
                   OnTop.Checked = False
          End If

          SkinForm1.SkinPath = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path")
          MediaPlayer1.ShowControls = False
          MediaPlayer1.ShowStatusBar = False
          CooLine1.Left = myReadINI((Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path")) + "\sflake_info.sfl", "label", "x", "") '+ 25
          CooLine1.Top = myReadINI((Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path")) + "\sflake_info.sfl", "label", "y", "") '+ 25
          CooLine1.Width = myReadINI((Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path")) + "\sflake_info.sfl", "label", "w", "")
          Frame12.Left = 10000
          Frame4.Left = 10000
          ctListBar1.Left = 10000
          Frame2.Top = myReadINI((Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path")) + "\sflake_info.sfl", "video", "y", "") '+ 25
          Frame2.Left = myReadINI((Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path")) + "\sflake_info.sfl", "video", "x", "") '+ 25
          Frame2.Width = myReadINI((Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path")) + "\sflake_info.sfl", "video", "w", "")
          Frame2.Height = myReadINI((Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path")) + "\sflake_info.sfl", "video", "h", "")
          Me.Width = myReadINI((Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path")) + "\sflake_info.sfl", "form", "w", "")
          Me.Height = myReadINI((Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path")) + "\sflake_info.sfl", "form", "h", "")
          
          Frame5.Left = 10000
          Frame6.Left = 10000
          Frame7.Left = 10000
          MediaPlayer1.Top = 0
          MediaPlayer1.Top = 0
          MediaPlayer1.Width = Frame2.Width
          MediaPlayer1.Height = Frame2.Height
          Image10.Picture = LoadPicture(Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path") + "\sflake_logo.bmp")
          Snow.Checked = True
          More.Checked = False
          If Len(RA.Source) > 0 And RA.GetPosition > 0 Then Call BigSma
                   Case 2
                  If f200 = 1 Then
                    MsgBox ("视频窗口已经打开无需再次打开.")
                    Exit Sub
                    End If
        If Len(MediaPlayer1.FileName) > 0 And MediaPlayer1.ImageSourceWidth > 0 Then
        Form200.Show
        jn = MediaPlayer1.FileName
        jd = MediaPlayer1.CurrentPosition
        MediaPlayer1.FileName = "LLXX"
        Form200.MediaPlayer1.FileName = jn
       Form200.MediaPlayer1.CurrentPosition = jd
       
        f200 = 1
        Else:
        If Len(RA.Source) > 0 And RA.GetPosition > 0 And RA.GetClipWidth > 0 Then
         Form200.Show
        jn = RA.Source
        jd = RA.GetPosition
        RMStop = False
         RA.DoStop
        Form200.RA.Source = jn
       Form200.RA.SetPosition jd
       f200 = 1
        Else
        MsgBox ("无可用视频,请先选定视频或图片文件再打开视频窗口.")
        End If
        End If
        Case 3
        If Len(MediaPlayer1.FileName) > 0 And MediaPlayer1.ImageSourceWidth > 0 Then
            Me.LyfTools1.MakeTop Me, False
            OnTop.Checked = False
        MediaPlayer1.DisplaySize = mpFullScreen
          Else:
            If Len(RA.Source) > 0 And RA.GetPosition > 0 And RA.GetClipWidth > 0 Then
                  Me.LyfTools1.MakeTop Me, False
                    OnTop.Checked = False
                   RA.SetFullScreen
             Else
          MsgBox ("无可用视频,请先选定视频或图片文件再开始全屏欣赏.")
           End If
          End If
        Case 4
        If Len(MediaPlayer1.FileName) > 0 Or (Len(RA.Source) > 0 And RA.GetPosition > 0) Then
        If Label26.Caption = "locked" Then
        Label26.Caption = "unlock"
        MsgBox ("已经解除对以下曲目的锁定: " & ListFile.List(pid) & " .")
         Exit Sub
        End If
        If Label26.Caption = "unlock" Then
        Label26.Caption = "locked"
        Check1.Value = 0
        Check2.Value = 0
              Call ReMem
        MsgBox ("已经锁定播放以下曲目: " & ListFile.List(pid) & " .")
        End If
        Else: MsgBox ("还没选择要锁定的曲目,请选定好曲目后再锁定.")
        End If
            Case 5
            If Len(MediaPlayer1.FileName) > 0 Or (Len(RA.Source) > 0 And RA.GetPosition > 0) Then
            Form5.Show
            Else
            MsgBox ("当前没有正在播放的媒体,无法标记媒体书签.请在正在播放媒体文件时再作尝试.")
            End If
            Case 6
            If Len(MediaPlayer1.FileName) = 0 Then
                     MsgBox ("找不到可以显示字幕的媒体文件,请确认该媒体含有字幕部分后重试.")
                If MediaPlayer1.ShowCaptioning = True Then
                  MediaPlayer1.ShowCaptioning = False
                  Image10.Height = 3375
                  Image10.Stretch = False
                   End If
                     Exit Sub
                End If
               If MediaPlayer1.ShowCaptioning = True Then
                  MediaPlayer1.ShowCaptioning = False
                  Image10.Height = 3375
                  Image10.Stretch = False
                 Exit Sub
                   End If
                  If MediaPlayer1.ShowCaptioning = False Then
                   MediaPlayer1.ShowCaptioning = True
                       Image10.Height = 1860
                       Image10.Stretch = True
                      Exit Sub
                      End If
         Case 7
          Dim fName As String
      If Len(MediaPlayer1.FileName) > 0 Or (Len(RA.Source) > 0 And RA.GetPosition > 0) Then
     If Len(MediaPlayer1.FileName) > 0 Then
       fName = MediaPlayer1.FileName
       Else
        fName = RA.GetSource
     End If
         Me.LyfTools1.ShowProp fName, Me
        Else: MsgBox ("没有找到可以显示媒体属性的媒体文件,请先选择好媒体文件再要求显示媒体属性.")
        End If
        Case 8
        On Error Resume Next
        If Len(MediaPlayer1.FileName) > 0 Then
         MediaPlayer1.ShowDialog mpShowDialogStatistics
         Exit Sub
       Else
       If Len(RA.Source) > 0 And RA.GetPosition > 0 Then
       RA.SetShowStatistics True
        Exit Sub
        End If
        End If
        MsgBox ("没有找到可以显示统计信息的媒体项目,请确认正在播放媒体文件再要求显示统计信息.")
        End Select
      
        
        
        
        
        
        
        
        
        
        
        
        
        Case 3
            Select Case nItem
            Case 1
            Dim idc As Integer
            pid = -1
           CommonDialog1.Filter = "媒体文件:Cda、Mp3、Wma、Wmv、Wav、Wax、Ra、Rm、Asf、Rmi、Asx、Mov、M1v、Mp2、Mpg、Mpeg、Mpa、Mpe、Avi、Mid、Qt、Aif、Aifc、Aiff、Au、Snd、Smi、Smil、Rt、Mpv、Rp、Ram、Rmm、Rtx" & _
          "|*.cda;*.au;*.and;*.aif;*.wmv;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.wma;*.wav;*.avi;*.mid;*.smi;*.smil;*.rt;*.mpv;*.rp;*.ram;*.rmm;*.rtx;*.ra;*.rm|所有文件:*.*|*.*"
          CommonDialog1.FilterIndex = 1
          CommonDialog1.FileName = ""
          CommonDialog1.ShowOpen
          If Len(CommonDialog1.FileName) > 0 Then
            For idc = 0 To ListFile.ListCount - 1
            If ListFile.List(idc) = CommonDialog1.FileName Then
            pid = idc
            ListFile.ListIndex = pid
            Exit For
            End If
            Next
            If pid = -1 Then
          ListFile.AddItem (CommonDialog1.FileName)
          pid = ListFile.ListCount - 1
          ListFile.ListIndex = pid
          End If
            If Right(ListFile.List(pid), 4) = ".cda" Then
    Call Image17_Click
    If (CDLoad) Then
        If (Track <= TotalTracks) Then
            If (Playing) Then
                Command = "play cd from " & Val(Left(Right(ListFile.List(pid), 6), 2))
                SendMCIString Command, True
             Else
                Command = "seek cd to " & Val(Left(Right(ListFile.List(pid), 6), 2))
                SendMCIString Command, True
                SendMCIString "play cd", True
                Playing = True
            End If
        End If
       Else
        SendMCIString "seek cd to 1", True
    End If
    Update
    If Len(TrackTime.Caption) > 0 And Playing = True Then
If RA.Left = 0 Then RMStop = False
If Len(MediaPlayer1.FileName) > 0 Then MediaPlayer1.FileName = "LLXX"
If RA.Left = 0 Then RMStop = False
RA.DoStop
RA.Left = 10000
SLD1.Left = 10000
End If


  Exit Sub
  End If

         MediaPlayer1.FileName = ListFile.List(pid)
           Label1.Caption = "media"
           TrackSelection.Left = 10000
           Frame8.Left = 2880
           Frame1.Left = 10000
          End If
            
            
            
            
            
            
            
            
            
            
            
        Case 2
         Me.MousePointer = 11
    If FileExists(Label27.Caption + "MPEGAV\") = True Then
          File1.Path = Label27.Caption + "MPEGAV\"
                 File1.Pattern = "*.dat;*.au;*.and;*.aif;*.wmv;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.wma;*.wav;*.avi;*.mid;*.smi;*.smil;*.rt;*.mpv;*.rp;*.ram;*.rmm;*.rtx;*.ra;*.rm"
          Dim ind As Integer
              For ind = 0 To File1.ListCount - 1
             ListFile.AddItem Label27.Caption + "MPEGAV\" + File1.List(ind), ListFile.ListCount
                  Next ind
            pid = ListFile.ListCount - File1.ListCount
                    MediaPlayer1.FileName = ListFile.List(pid)
        Label1.Caption = "media"
        TrackSelection.Left = 10000
       Frame8.Left = 2880
        Frame1.Left = 10000
        Else: MsgBox ("光盘中没有VCD光盘,请先放入VCD光盘再使用本功能.")
        End If
         Me.MousePointer = 0
               Case 3
                Me.MousePointer = 11
               Dim firstpath As String, NumFiles, dircount As Integer
               Dir1.Path = Label27.Caption
               File1.Pattern = "*.au;*.and;*.aif;*.wmv;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.wma;*.wav;*.avi;*.mid;*.smi;*.smil;*.rt;*.mpv;*.rp;*.ram;*.rmm;*.rtx;*.ra;*.rm"
               If File1.Path = Label3.Caption Then GoTo cc
                 firstpath = Dir1.Path
                dircount = Dir1.ListCount
                 NumFiles = 0                       ' Reset global foundfiles indicator.
               Result = DirDiver(firstpath, dircount, "")
                  File1.Path = Dir1.Path
                       If Cdno = 0 Then
cc:
                     MsgBox ("在光驱中找不到含有媒体文件的光盘.请放入含有媒体文件的光盘后重试.")
                  Me.MousePointer = 0
                     Exit Sub
                     End If
                    pid = ListFile.ListCount - Cdno
              MediaPlayer1.FileName = ListFile.List(pid)
            Cdno = 0
                Label1.Caption = "media"
        TrackSelection.Left = 10000
       Frame8.Left = 2880
        Frame1.Left = 10000
         Me.MousePointer = 0
   Case 4
  If Me.FileExists(Label3.Caption + "\SmM_FP.exe") = True Then
         Shell Label3.Caption + "\SmM_FP.exe", vbNormalFocus
       Else
         MsgBox ("找不到文件[ " + Label3.Caption + "\SmM_FP.exe" + " ].该文件可能已经丢失或被移动,请重新安装 Snowman Media ilxz 3.5")
       End If
   End Select
Case 5
   Select Case nItem
     Case 1
                 If Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name") <> "Error" Then
          pid = -1
          
                  For idc = 0 To ListFile.ListCount
            If ListFile.List(idc) = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name") Then
            pid = idc
            ListFile.ListIndex = pid
               MediaPlayer1.FileName = ListFile.List(pid)
       MediaPlayer1.CurrentPosition = Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute")
            If Len(MediaPlayer1.FileName) = 0 Then RA.SetPosition Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute")

            End If
            Next
          If pid = -1 Then
          ListFile.AddItem (Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name"))
          pid = ListFile.ListCount - 1
          ListFile.ListIndex = pid
           MediaPlayer1.FileName = ListFile.List(pid)
       MediaPlayer1.CurrentPosition = Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute")
            If Len(MediaPlayer1.FileName) = 0 Then RA.SetPosition Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute")
          End If
         Else: MsgBox ("断点记录为空,没有可以用于播放的媒体文件记录.请在记录断点信息后再次尝试.")
         End If
    
          Case 2
                      If Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_A") <> "Error" Then
          pid = -1
          
                  For idc = 0 To ListFile.ListCount
            If ListFile.List(idc) = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_A") Then
            pid = idc
            ListFile.ListIndex = pid
               MediaPlayer1.FileName = ListFile.List(pid)
       MediaPlayer1.CurrentPosition = Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_A")
            If Len(MediaPlayer1.FileName) = 0 Then RA.SetPosition Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_A")

            End If
            Next
          If pid = -1 Then
          ListFile.AddItem (Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_A"))
          pid = ListFile.ListCount - 1
          ListFile.ListIndex = pid
           MediaPlayer1.FileName = ListFile.List(pid)
       MediaPlayer1.CurrentPosition = Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_A")
            If Len(MediaPlayer1.FileName) = 0 Then RA.SetPosition Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_a")
          End If
         Else: MsgBox ("本媒体书签为空,没有可以用于播放的媒体文件记录.请在标记本书签后再次尝试.")
         End If
    
        Case 3
                     If Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_B") <> "Error" Then
          pid = -1
          
                  For idc = 0 To ListFile.ListCount
            If ListFile.List(idc) = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_B") Then
            pid = idc
            ListFile.ListIndex = pid
               MediaPlayer1.FileName = ListFile.List(pid)
       MediaPlayer1.CurrentPosition = Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_B")
            If Len(MediaPlayer1.FileName) = 0 Then RA.SetPosition Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_B")

            End If
            Next
          If pid = -1 Then
          ListFile.AddItem (Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_B"))
          pid = ListFile.ListCount - 1
          ListFile.ListIndex = pid
           MediaPlayer1.FileName = ListFile.List(pid)
       MediaPlayer1.CurrentPosition = Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_B")
            If Len(MediaPlayer1.FileName) = 0 Then RA.SetPosition Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_B")
          End If
         Else: MsgBox ("本媒体书签为空,没有可以用于播放的媒体文件记录.请在标记本书签后再次尝试.")
         End If
         Case 4
                      If Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_C") <> "Error" Then
          pid = -1
          
                  For idc = 0 To ListFile.ListCount
            If ListFile.List(idc) = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_C") Then
            pid = idc
            ListFile.ListIndex = pid
               MediaPlayer1.FileName = ListFile.List(pid)
       MediaPlayer1.CurrentPosition = Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_C")
            If Len(MediaPlayer1.FileName) = 0 Then RA.SetPosition Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_C")

            End If
            Next
          If pid = -1 Then
          ListFile.AddItem (Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_C"))
          pid = ListFile.ListCount - 1
          ListFile.ListIndex = pid
           MediaPlayer1.FileName = ListFile.List(pid)
       MediaPlayer1.CurrentPosition = Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_C")
            If Len(MediaPlayer1.FileName) = 0 Then RA.SetPosition Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_C")
          End If
         Else: MsgBox ("本媒体书签为空,没有可以用于播放的媒体文件记录.请在标记本书签后再次尝试.")
         End If
       Case 5
                   If Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_D") <> "Error" Then
          pid = -1
          
                  For idc = 0 To ListFile.ListCount
            If ListFile.List(idc) = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_D") Then
            pid = idc
            ListFile.ListIndex = pid
               MediaPlayer1.FileName = ListFile.List(pid)
       MediaPlayer1.CurrentPosition = Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_D")
            If Len(MediaPlayer1.FileName) = 0 Then RA.SetPosition Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_D")

            End If
            Next
          If pid = -1 Then
          ListFile.AddItem (Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_D"))
          pid = ListFile.ListCount - 1
          ListFile.ListIndex = pid
           MediaPlayer1.FileName = ListFile.List(pid)
       MediaPlayer1.CurrentPosition = Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_D")
            If Len(MediaPlayer1.FileName) = 0 Then RA.SetPosition Form102.LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_D")
          End If
         Else: MsgBox ("本媒体书签为空,没有可以用于播放的媒体文件记录.请在标记本书签后再次尝试.")
         End If
        
         
         
         
         
         
         
         
         
         
       
     End Select
       
     Case 2
          Select Case nItem
             Case 1
          CommonDialog1.Filter = "媒体文件:Mp3、Wma、Wmv、Wav、Wax、Ra、Rm、Asf、Rmi、Asx、Mov、M1v、Mp2、Mpg、Mpeg、Mpa、Mpe、Avi、Mid、Qt、Aif、Aifc、Aiff、Au、Snd、Smi、Smil、Rt、Mpv、Rp、Ram、Rmm、Rtx" & _
          "|*.au;*.and;*.aif;*.wmv;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.wma;*.wav;*.avi;*.mid;*.smi;*.smil;*.rt;*.mpv;*.rp;*.ram;*.rmm;*.rtx;*.ra;*.rm|所有文件:*.*|*.*"
          CommonDialog1.FilterIndex = 1
          CommonDialog1.FileName = ""
          CommonDialog1.ShowOpen
          If Len(CommonDialog1.FileName) > 0 Then
       SelectFileName = CommonDialog1.FileName
       If ListFile.ListCount > 0 Then
      ListFile.AddItem SelectFileName, ListFile.ListCount
     Else
      ListFile.AddItem SelectFileName, 0
     End If
      Label1.Caption = "media"
      TrackSelection.Left = 10000
 Frame8.Left = 2880
 Frame1.Left = 10000
     End If
     Case 2
       Dim AddUrl As String
       AddUrl = InputBox("请输入你要加入到列表的媒体文件的 URL 地址.地址可以是 Internet 上的也可以是本地主机的,Snowman Media ilxz 3.5 将自动识别并进行播放.")
       If Len(AddUrl) > 0 Then
        ListFile.AddItem AddUrl
      End If
     Case 3
     If Len(ListFile.List(ListFile.ListIndex)) = 0 Then
           MsgBox ("还没有选定要删除的曲目,请先选定在后再删除.")
           Exit Sub
           End If
         If ListFile.ListCount > 0 And ListFile.SelCount > 0 Then
            If pid <= ListFile.ListIndex Then pid = pid - 1
          ListFile.RemoveItem (ListFile.ListIndex)
        
              End If
                Case 5
                CommonDialog1.FileName = ""
  CommonDialog1.Filter = "列表文件:M3u" & _
          "|*.m3u|所有文件:*.*|*.*"
             CommonDialog1.ShowOpen
             If Len(CommonDialog1.FileName) > 0 Then
              Open CommonDialog1.FileName For Input As #1
           While Not EOF(1)
          Line Input #1, test
           ListFile.AddItem RTrim(test)
           Wend
             Close #1
            End If
            Case 4
             Me.MousePointer = 11
             Dim BI As BROWSEINFO
             Dim nFolder As Long
             Dim IDL As ITEMIDLIST
             Dim pIdl As Long
             Dim sPath As String
             Dim SHFI As SHFILEINFO
             Dim m_wCurOptIdx As Integer
             Dim txtPath As String
             Dim txtDisplayName As String
             Dim noerror, SHGFI_PIDL, SHGFI_ICON, SHGFI_SMALLICON As Integer
             With BI
              .hOwner = Me.hwnd
            nFolder = GetFolderValue(m_wCurOptIdx)
            If SHGetSpecialFolderLocation(ByVal Me.hwnd, ByVal nFolder, IDL) = noerror Then
            .pidlRoot = IDL.mkid.cb
             End If
     .pszDisplayName = String$(MAX_PATH, 0)
    .lpszTitle = "请选择你要添加媒体文件夹.Snowman Media ilxz 3.5 将自动把文件夹内的媒体加入列表."
    .ulFlags = 0
                  End With
         txtPath = ""
           txtDisplayName = ""
             pIdl = SHBrowseForFolder(BI)
              If pIdl = 0 Then
              Me.MousePointer = 0
              Exit Sub
              End If
               sPath = String$(MAX_PATH, 0)
                    SHGetPathFromIDList ByVal pIdl, ByVal sPath
                 txtPath = Left(sPath, InStr(sPath, vbNullChar) - 1)
                    txtDisplayName = Left$(BI.pszDisplayName, _
                    InStr(BI.pszDisplayName, vbNullChar) - 1)
                 SHGetFileInfo ByVal pIdl, 0&, SHFI, Len(SHFI), _
                SHGFI_PIDL Or SHGFI_ICON Or SHGFI_SMALLICON
                SHGetFileInfo ByVal pIdl, 0&, SHFI, Len(SHFI), _
                SHGFI_PIDL Or SHGFI_ICON
                    CoTaskMemFree pIdl
               
                 Dir1.Path = txtPath
               File1.Pattern = "*.au;*.and;*.aif;*.wmv;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.wma;*.wav;*.avi;*.mid;*.smi;*.smil;*.rt;*.mpv;*.rp;*.ram;*.rmm;*.rtx;*.ra;*.rm"
                 firstpath = Dir1.Path
                dircount = Dir1.ListCount
                 NumFiles = 0                       ' Reset global foundfiles indicator.
               Result = DirDiver(firstpath, dircount, "")
                  File1.Path = Dir1.Path
          Me.MousePointer = 0
    Case 9
                ListFile.Clear
                  Case 8
            If ListFile.ListCount > 0 Then
                  Dim i As Integer
              CommonDialog1.FileName = ""
  CommonDialog1.Filter = "列表文件:M3u" & _
          "|*.m3u|所有文件:*.*|*.*"
             CommonDialog1.ShowSave
             If Len(CommonDialog1.FileName) > 0 Then
                  If Right(CommonDialog1.FileName, 4) <> ".m3u" Then CommonDialog1.FileName = CommonDialog1.FileName + ".m3u"
             Open CommonDialog1.FileName For Output As #1
    For i = 0 To ListFile.ListCount - 1
     Print #1, ListFile.List(i)
    Next i
   Close (1)
   End If
        Else: MsgBox ("当前播放列表为空无法保存,请先加入曲目到列表.")
        End If
      Case 7
      Call Tiday
       Case 6
            pid = -1
            'MediaPlayer1.Filename = ListFile.List(0)
            'pid = 1
            MdNo = 10
            MdBo = True
    End Select
     
     
     
     
     
     
     
     
     Case 4
      Select Case nItem
        
    
        Case 1
         If Form102.LyfTools1.IsConnected Then
             Label2.Caption = "media"
             Formo.Show
             Else
               Label2.Caption = "media"
            Form2.Show
            End If
          Case 2
           If Form102.LyfTools1.IsConnected Then
          Label2.Caption = "flash"
             Formo.Show
                  Else
                      Label2.Caption = "flash"
                  Form2.Show
            End If
            End Select




     Case 6
      Select Case nItem
        Case 1
          File1.Pattern = "*.au;*.and;*.aif;*.wmv;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.wma;*.wav;*.avi;*.mid;*.smi;*.smil;*.rt;*.mpv;*.rp;*.ram;*.rmm;*.rtx;*.ra;*.rm"
         Call ToolB
          Case 2
       Form100.Show
       Case 3
      If Me.FileExists(Label3.Caption + "\SmM_Od.exe") = True Then
         Shell Label3.Caption + "\SmM_Od.exe", vbNormalFocus
       Else
         MsgBox ("找不到 Snowman Media ilxz 3.5 功能插件管理器,请正确安装 Snowman Media ilxz 3.5 功能插件管理器后重试.")
       End If
       Case 4
         Form4.Show
        Case 5
           MediaPlayer1.ShowDialog mpShowDialogOptions
        Case 6
  If Me.FileExists(Label3.Caption + "\SmM_St.exe") = True Then
         Shell Label3.Caption + "\SmM_St.exe", vbNormalFocus
         
         
       Else
         MsgBox ("找不到文件[ " + Label3.Caption + "\SmM_St.exe" + " ].该文件可能已经丢失或被移动,请重新安装 Snowman Media ilxz 3.5")
       End If
        Case 7
  If Me.FileExists(Me.LyfTools1.GetWinPath + "\SNDVOL32.exe") = True Then
         Shell Me.LyfTools1.GetWinPath + "\SNDVOL32.exe", vbNormalFocus
         
         
       Else
         MsgBox ("找不到文件[ " + Me.LyfTools1.GetWinPath + "\SNDVOL32.exe" + " ].该系统文件可能已经丢失或无效")
       End If
            End Select

    Case 7
      Select Case nItem
        Case 1
  If Me.FileExists(Label3.Caption + "\SmM_Hp.exe") = True Then
         Shell Label3.Caption + "\SmM_Hp.exe", vbNormalFocus
       Else
         MsgBox ("找不到文件[ " + Label3.Caption + "\SmM_Hp.exe" + " ].该文件可能已经丢失或被移动,请重新安装 Snowman Media ilxz 3.5")
       End If
          Case 2
  If Me.FileExists(Label3.Caption + "\SmM_Ud.exe") = True Then
         Shell Label3.Caption + "\SmM_Ud.exe", vbNormalFocus
       Else
         MsgBox ("找不到文件[ " + Label3.Caption + "\SmM_Ud.exe" + " ].该文件可能已经丢失或被移动,请重新安装 Snowman Media ilxz 3.5")
       End If
        Case 3
         frmAbout.Show
   Case 4
         If Me.FileExists(Label3.Caption + "\自述文件.txt") = True Then
         
        Me.LyfTools1.OpenFile (Label3.Caption + "\自述文件.txt")
       Else
         MsgBox ("找不到文件[ " + Label3.Caption + "\自述文件.txt" + " ].该文件可能已经丢失或被移动,请重新安装 Snowman Media ilxz 3.5")
       End If

   Case 5
                If Me.FileExists(Label3.Caption + "\许可协议.txt") = True Then
         Me.LyfTools1.OpenFile (Label3.Caption + "\许可协议.txt")
       Else
         MsgBox ("找不到文件[ " + Label3.Caption + "\许可协议.txt" + " ].该文件可能已经丢失或被移动,请重新安装 Snowman Media ilxz 3.5")
       End If

 
        Case 6
              If Form102.LyfTools1.IsConnected = True Then
              
              Form102.LyfTools1.HttpTo ("http://www.h2ont.com")
              Else
               MsgBox ("无法访问 流动网络 可能你的计算机尚未连接网络,请确认连接网络后重试.")
               End If
         Case 7
               Form102.LyfTools1.SendMail ("leask@21cn.com")
          Case 8
              Unload Me
              End

End Select
End Select
End Sub

Private Sub Window_Click()
Call ctListBar1_ItemClick(1, 2)
End Sub

Private Sub Xh_Click()
Call ctListBar1_ItemClick(6, 6)
End Sub
