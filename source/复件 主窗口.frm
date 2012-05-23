VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "SmM_Snowflake.ocx"
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "lyftools.ocx"
Object = "{244E6785-6684-11D2-943F-A976CFB4FC0C}#1.0#0"; "ctlstbar.ocx"
Object = "{C40E7B9F-6CF0-11D2-AA70-444553540000}#1.0#0"; "Coolineprj.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.OCX"
Object = "{38943DFD-2C76-11D5-8FCF-A3B833033124}#1.0#0"; "SmM_SouCtrl.ocx"
Object = "{33155A3D-0CE0-11D1-A6B4-444553540000}#1.0#0"; "SysTray.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Snowman Media ilxz"
   ClientHeight    =   4245
   ClientLeft      =   1230
   ClientTop       =   1590
   ClientWidth     =   7335
   DrawStyle       =   6  'Inside Solid
   Icon            =   "主窗口.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4245
   ScaleWidth      =   7335
   Begin SysTray.SystemTray ST1 
      Left            =   3780
      Top             =   7110
      _ExtentX        =   847
      _ExtentY        =   847
      SysTrayText     =   "Snowman Media ilxz"
      IconFile        =   0
   End
   Begin HYZ声音控制控件.HYZVolBan HB 
      Height          =   330
      Left            =   5040
      TabIndex        =   15
      Top             =   7200
      Visible         =   0   'False
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   582
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   510
      Left            =   3015
      TabIndex        =   14
      Top             =   7110
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   5625
      TabIndex        =   13
      Top             =   7155
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox TrackSelection 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   300
      ItemData        =   "主窗口.frx":2372
      Left            =   6165
      List            =   "主窗口.frx":2374
      MouseIcon       =   "主窗口.frx":2376
      TabIndex        =   12
      Text            =   "S.M."
      ToolTipText     =   "曲目"
      Top             =   7200
      Visible         =   0   'False
      Width           =   435
   End
   Begin API控制大全.LyfTools Ly 
      Left            =   2385
      Top             =   7110
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4455
      Top             =   7110
   End
   Begin VB.Frame Fm0 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   5865
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   0
      Width           =   8055
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   8
         ToolTipText     =   "状态"
         Top             =   4365
         Width           =   4245
         Begin VB.TextBox TimeWindow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   90
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "[00]00:00"
            ToolTipText     =   "播放时间"
            Top             =   45
            Width           =   1185
         End
         Begin VB.Label TrackTime 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   1710
            OLEDropMode     =   1  'Manual
            TabIndex        =   11
            ToolTipText     =   "曲目时间"
            Top             =   45
            Width           =   90
         End
         Begin VB.Label TotalTrack 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   2970
            OLEDropMode     =   1  'Manual
            TabIndex        =   10
            ToolTipText     =   "总时间"
            Top             =   45
            Width           =   90
         End
      End
      Begin VB.Frame Fm5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C08062&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   1635
         Left            =   4500
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Top             =   3690
         Width           =   105
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            Height          =   1125
            Left            =   0
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口.frx":24C8
            Top             =   0
            Width           =   150
         End
      End
      Begin CTLISTBARLibCtl.ctListBar cLT1 
         Height          =   1230
         Left            =   4635
         TabIndex        =   4
         ToolTipText     =   "功能菜单"
         Top             =   3645
         Width           =   2760
         _Version        =   65536
         _ExtentX        =   4868
         _ExtentY        =   2170
         _StockProps     =   70
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
         BackImage       =   "主窗口.frx":254F
         BorderColor     =   12615778
         ButtonBackColor =   16777215
         ButtonForeColor =   12615778
         ListBackColor   =   12615778
         ListForeColor   =   16711680
         BarBackColor    =   12615748
         BarForeColor    =   12615778
         BorderType      =   1
         ListBarStyle    =   0
         BarHeight       =   0
         WordWrap        =   -1  'True
         Caption         =   ""
         PicArray0       =   "主窗口.frx":3173
         PicArray1       =   "主窗口.frx":52AD
         PicArray2       =   "主窗口.frx":5B87
         PicArray3       =   "主窗口.frx":BE21
         PicArray4       =   "主窗口.frx":DF5B
         PicArray5       =   "主窗口.frx":10C65
         PicArray6       =   "主窗口.frx":132AF
         PicArray7       =   "主窗口.frx":14101
         PicArray8       =   "主窗口.frx":1A073
         PicArray9       =   "主窗口.frx":2030D
         PicArray10      =   "主窗口.frx":20329
         PicArray11      =   "主窗口.frx":20345
         PicArray12      =   "主窗口.frx":20361
         PicArray13      =   "主窗口.frx":2037D
         PicArray14      =   "主窗口.frx":20399
         PicArray15      =   "主窗口.frx":203B5
         PicArray16      =   "主窗口.frx":203D1
         PicArray17      =   "主窗口.frx":203ED
         PicArray18      =   "主窗口.frx":20409
         PicArray19      =   "主窗口.frx":20425
         PicArray20      =   "主窗口.frx":20441
         PicArray21      =   "主窗口.frx":2045D
         PicArray22      =   "主窗口.frx":20479
         PicArray23      =   "主窗口.frx":20495
         PicArray24      =   "主窗口.frx":204B1
         PicArray25      =   "主窗口.frx":204CD
         PicArray26      =   "主窗口.frx":204E9
         PicArray27      =   "主窗口.frx":20505
         PicArray28      =   "主窗口.frx":20521
         PicArray29      =   "主窗口.frx":2053D
         PicArray30      =   "主窗口.frx":20559
         PicArray31      =   "主窗口.frx":20575
         PicArray32      =   "主窗口.frx":20591
         PicArray33      =   "主窗口.frx":205AD
         PicArray34      =   "主窗口.frx":205C9
         PicArray35      =   "主窗口.frx":205E5
         PicArray36      =   "主窗口.frx":20601
         PicArray37      =   "主窗口.frx":2061D
         PicArray38      =   "主窗口.frx":20639
         PicArray39      =   "主窗口.frx":20655
         PicArray40      =   "主窗口.frx":20671
         PicArray41      =   "主窗口.frx":2068D
         PicArray42      =   "主窗口.frx":206A9
         PicArray43      =   "主窗口.frx":206C5
         PicArray44      =   "主窗口.frx":206E1
         PicArray45      =   "主窗口.frx":206FD
         PicArray46      =   "主窗口.frx":20719
         PicArray47      =   "主窗口.frx":20735
         PicArray48      =   "主窗口.frx":20751
         PicArray49      =   "主窗口.frx":2076D
         PicArray50      =   "主窗口.frx":20789
         PicArray51      =   "主窗口.frx":207A5
         PicArray52      =   "主窗口.frx":207C1
         PicArray53      =   "主窗口.frx":207DD
         PicArray54      =   "主窗口.frx":207F9
         PicArray55      =   "主窗口.frx":20815
         PicArray56      =   "主窗口.frx":20831
         PicArray57      =   "主窗口.frx":2084D
         PicArray58      =   "主窗口.frx":20869
         PicArray59      =   "主窗口.frx":20885
         PicArray60      =   "主窗口.frx":208A1
         PicArray61      =   "主窗口.frx":208BD
         PicArray62      =   "主窗口.frx":208D9
         PicArray63      =   "主窗口.frx":208F5
         PicArray64      =   "主窗口.frx":20911
         PicArray65      =   "主窗口.frx":2092D
         PicArray66      =   "主窗口.frx":20949
         PicArray67      =   "主窗口.frx":20965
         PicArray68      =   "主窗口.frx":20981
         PicArray69      =   "主窗口.frx":2099D
         PicArray70      =   "主窗口.frx":209B9
         PicArray71      =   "主窗口.frx":209D5
         PicArray72      =   "主窗口.frx":209F1
         PicArray73      =   "主窗口.frx":20A0D
         PicArray74      =   "主窗口.frx":20A29
         PicArray75      =   "主窗口.frx":20A45
         PicArray76      =   "主窗口.frx":20A61
         PicArray77      =   "主窗口.frx":20A7D
         PicArray78      =   "主窗口.frx":20A99
         PicArray79      =   "主窗口.frx":20AB5
         PicArray80      =   "主窗口.frx":20AD1
         PicArray81      =   "主窗口.frx":20AED
         PicArray82      =   "主窗口.frx":20B09
         PicArray83      =   "主窗口.frx":20B25
         PicArray84      =   "主窗口.frx":20B41
         PicArray85      =   "主窗口.frx":20B5D
         PicArray86      =   "主窗口.frx":20B79
         PicArray87      =   "主窗口.frx":20B95
         PicArray88      =   "主窗口.frx":20BB1
         PicArray89      =   "主窗口.frx":20BCD
         PicArray90      =   "主窗口.frx":20BE9
         PicArray91      =   "主窗口.frx":20C05
         PicArray92      =   "主窗口.frx":20C21
         PicArray93      =   "主窗口.frx":20C3D
         PicArray94      =   "主窗口.frx":20C59
         PicArray95      =   "主窗口.frx":20C75
         PicArray96      =   "主窗口.frx":20C91
         PicArray97      =   "主窗口.frx":20CAD
         PicArray98      =   "主窗口.frx":20CC9
         PicArray99      =   "主窗口.frx":20CE5
      End
      Begin VB.Frame Fm10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C08044&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   7
         Top             =   3915
         Width           =   6990
         Begin VB.Image Image10 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2070
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口.frx":20D01
            Stretch         =   -1  'True
            ToolTipText     =   "弹出光盘"
            Top             =   0
            Width           =   375
         End
         Begin VB.Image Image3 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   405
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口.frx":21042
            Stretch         =   -1  'True
            ToolTipText     =   "倒退 5 秒"
            Top             =   0
            Width           =   330
         End
         Begin VB.Image Image9 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   3240
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口.frx":21383
            Stretch         =   -1  'True
            ToolTipText     =   "音量"
            Top             =   0
            Width           =   375
         End
         Begin VB.Image Image8 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2835
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口.frx":216C4
            Stretch         =   -1  'True
            ToolTipText     =   "下一首曲目"
            Top             =   0
            Width           =   375
         End
         Begin VB.Image Image7 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   2475
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口.frx":21A05
            Stretch         =   -1  'True
            ToolTipText     =   "块进 5 秒"
            Top             =   0
            Width           =   330
         End
         Begin VB.Image Image6 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   1620
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口.frx":21D46
            Stretch         =   -1  'True
            ToolTipText     =   "停止"
            Top             =   0
            Width           =   420
         End
         Begin VB.Image Image5 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   1260
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口.frx":22087
            Stretch         =   -1  'True
            ToolTipText     =   "暂停"
            Top             =   0
            Width           =   375
         End
         Begin VB.Image Image4 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   765
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口.frx":223C8
            Stretch         =   -1  'True
            ToolTipText     =   "播放"
            Top             =   0
            Width           =   465
         End
         Begin VB.Image Image2 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   45
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口.frx":22709
            Stretch         =   -1  'True
            ToolTipText     =   "上一首曲目"
            Top             =   0
            Width           =   330
         End
         Begin VB.Image Ig4 
            Appearance      =   0  'Flat
            Height          =   435
            Left            =   45
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口.frx":22A4A
            Top             =   0
            Width           =   3750
         End
      End
      Begin CooLinePrj.CooLine Cl0 
         Height          =   240
         Left            =   45
         TabIndex        =   3
         ToolTipText     =   "信息"
         Top             =   45
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   423
         InsChr          =   95
         Speed           =   120
         Display         =   "Enjoy your multimedia by using Snowman Media"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin VB.ListBox LF1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00C0FFFF&
         Height          =   2910
         IntegralHeight  =   0   'False
         ItemData        =   "主窗口.frx":23B99
         Left            =   4545
         List            =   "主窗口.frx":23B9B
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         ToolTipText     =   "曲目列表"
         Top             =   495
         Width           =   2670
      End
      Begin VB.Image Ig13 
         Appearance      =   0  'Flat
         Height          =   15
         Left            =   6705
         OLEDropMode     =   1  'Manual
         Picture         =   "主窗口.frx":23B9D
         ToolTipText     =   "不循环"
         Top             =   15
         Width           =   15
      End
      Begin VB.Image Ig12 
         Appearance      =   0  'Flat
         Height          =   15
         Left            =   6015
         OLEDropMode     =   1  'Manual
         Picture         =   "主窗口.frx":23BCC
         ToolTipText     =   "原序"
         Top             =   15
         Width           =   15
      End
      Begin VB.Image Ig11 
         Appearance      =   0  'Flat
         Height          =   15
         Left            =   5340
         OLEDropMode     =   1  'Manual
         Picture         =   "主窗口.frx":23BFB
         ToolTipText     =   "媒体"
         Top             =   15
         Width           =   15
      End
      Begin VB.Image Ig8 
         Appearance      =   0  'Flat
         Height          =   15
         Left            =   540
         OLEDropMode     =   1  'Manual
         Picture         =   "主窗口.frx":23C2A
         Top             =   765
         Width           =   15
      End
      Begin VB.Image Ig7 
         Appearance      =   0  'Flat
         Height          =   2610
         Left            =   45
         OLEDropMode     =   1  'Manual
         Picture         =   "主窗口.frx":23C59
         Stretch         =   -1  'True
         Top             =   405
         Width           =   3435
      End
      Begin VB.Image Ig2 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   0
         OLEDropMode     =   1  'Manual
         Picture         =   "主窗口.frx":23F9D
         Stretch         =   -1  'True
         Top             =   3330
         Width           =   7425
      End
      Begin VB.Image Ig0 
         Appearance      =   0  'Flat
         Height          =   510
         Left            =   0
         OLEDropMode     =   1  'Manual
         Picture         =   "主窗口.frx":243E7
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7665
      End
      Begin MediaPlayerCtl.MediaPlayer MP1 
         DragIcon        =   "主窗口.frx":2484B
         Height          =   4290
         Left            =   45
         TabIndex        =   5
         Top             =   405
         Width           =   4470
         AudioStream     =   -1
         AutoSize        =   0   'False
         AutoStart       =   -1  'True
         AnimationAtStart=   0   'False
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
         DisplayBackColor=   16777215
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
         ShowAudioControls=   0   'False
         ShowDisplay     =   0   'False
         ShowGotoBar     =   0   'False
         ShowPositionControls=   0   'False
         ShowStatusBar   =   -1  'True
         ShowTracker     =   -1  'True
         TransparentAtStart=   0   'False
         VideoBorderWidth=   0
         VideoBorderColor=   0
         VideoBorder3D   =   0   'False
         Volume          =   -440
         WindowlessVideo =   0   'False
      End
      Begin VB.Image Ig1 
         Appearance      =   0  'Flat
         Height          =   1785
         Left            =   0
         OLEDropMode     =   1  'Manual
         Picture         =   "主窗口.frx":2499D
         Stretch         =   -1  'True
         Top             =   3645
         Width           =   7590
      End
   End
   Begin ACTIVESKINLibCtl.SkinForm SF1 
      Height          =   480
      Left            =   1800
      OleObjectBlob   =   "主窗口.frx":24A30
      TabIndex        =   0
      Top             =   7110
      Width           =   480
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Snowman Media  3.0"
   End
   Begin VB.Menu sdgfregvcxv 
      Caption         =   "a"
      Begin VB.Menu sdfg 
         Caption         =   "Snowflake"
         Shortcut        =   %{BKSP}
      End
   End
   Begin VB.Menu a001 
      Caption         =   "b"
      Begin VB.Menu a00101 
         Caption         =   "浏览(&O)..."
         Shortcut        =   ^O
      End
      Begin VB.Menu a00102 
         Caption         =   "地址(&A)..."
         Shortcut        =   ^D
      End
      Begin VB.Menu a00103 
         Caption         =   "文件夹(&F)..."
         Shortcut        =   ^R
      End
      Begin VB.Menu a00104 
         Caption         =   "-"
      End
      Begin VB.Menu a00105 
         Caption         =   "CD (&C)"
      End
      Begin VB.Menu a00106 
         Caption         =   "VCD (&V)"
      End
      Begin VB.Menu a00107 
         Caption         =   "DVD (&D)"
      End
      Begin VB.Menu a00108 
         Caption         =   "媒体光盘(&L)"
      End
      Begin VB.Menu a00109 
         Caption         =   "-"
      End
      Begin VB.Menu ggg 
         Caption         =   "Flash (&S)"
      End
      Begin VB.Menu dsf 
         Caption         =   "-"
      End
      Begin VB.Menu a001010 
         Caption         =   "HDTV (&H)"
      End
      Begin VB.Menu a001011 
         Caption         =   "调谐收音机(&R)"
         Enabled         =   0   'False
      End
      Begin VB.Menu a001012 
         Caption         =   "-"
      End
      Begin VB.Menu sdtg56 
         Caption         =   "播放向导(&W)"
      End
      Begin VB.Menu a001013 
         Caption         =   "媒体指南(&M)"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu a002 
      Caption         =   "c"
      Begin VB.Menu grdfg 
         Caption         =   "播放(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu dfhgrtht 
         Caption         =   "暂停(&U)"
         Shortcut        =   ^U
      End
      Begin VB.Menu ytjuytjkuy 
         Caption         =   "停止(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu rtyu 
         Caption         =   "-"
      End
      Begin VB.Menu ytu5gf 
         Caption         =   "上一首曲目(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu fgrgfrgrgrtg 
         Caption         =   "下一首曲目(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu greg 
         Caption         =   "倒退 5 秒(&K)"
      End
      Begin VB.Menu dfgr 
         Caption         =   "快进 5 秒(&T)"
      End
      Begin VB.Menu gdsrg 
         Caption         =   "-"
      End
      Begin VB.Menu gdfg 
         Caption         =   "可视效果(&V)"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu ewtgreuyjk 
         Caption         =   "-"
      End
      Begin VB.Menu a00201 
         Caption         =   "字幕(&C)"
      End
      Begin VB.Menu a002020 
         Caption         =   "统计信息(&I)..."
         Shortcut        =   ^I
      End
      Begin VB.Menu a00202 
         Caption         =   "文件属性(&R)..."
      End
      Begin VB.Menu a00203 
         Caption         =   "-"
      End
      Begin VB.Menu a002045 
         Caption         =   "随机(&W)"
         Shortcut        =   ^W
      End
      Begin VB.Menu a00205 
         Caption         =   "循环(&L)"
         Shortcut        =   ^L
      End
      Begin VB.Menu fdhj 
         Caption         =   "-"
      End
      Begin VB.Menu fdhgrtjbvg 
         Caption         =   "谁与我同听一曲(&H)"
         Enabled         =   0   'False
      End
      Begin VB.Menu jjjyyy 
         Caption         =   "网络广播(&N)"
         Enabled         =   0   'False
      End
      Begin VB.Menu frgfrefaew 
         Caption         =   "-"
      End
      Begin VB.Menu dfg 
         Caption         =   "弹出(&E)"
         Shortcut        =   ^J
      End
   End
   Begin VB.Menu a003 
      Caption         =   "d"
      Begin VB.Menu a00301 
         Caption         =   "原始尺寸(&O)"
      End
      Begin VB.Menu a00302 
         Caption         =   "双倍尺寸(&D)"
      End
      Begin VB.Menu a00303 
         Caption         =   "最大化(&X)"
      End
      Begin VB.Menu a00304 
         Caption         =   "最小化(&N)"
      End
      Begin VB.Menu a00305 
         Caption         =   "-"
      End
      Begin VB.Menu a00308 
         Caption         =   "始终前置(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu fsdcfd555 
         Caption         =   "-"
      End
      Begin VB.Menu hgnmbmjh 
         Caption         =   "视频缩放(&R)"
         Begin VB.Menu fghgt 
            Caption         =   "窗口适应视频(&I)"
         End
         Begin VB.Menu dfxswefrff 
            Caption         =   "-"
         End
         Begin VB.Menu mjjhjm 
            Caption         =   "视频适应窗口(&A)"
         End
         Begin VB.Menu xcdcdcd 
            Caption         =   "-"
         End
         Begin VB.Menu cerde 
            Caption         =   "50% (&L)"
         End
         Begin VB.Menu cbcxbf 
            Caption         =   "100% (&R)"
         End
         Begin VB.Menu xfg 
            Caption         =   "200% (&B)"
         End
         Begin VB.Menu xbf 
            Caption         =   "-"
         End
         Begin VB.Menu a00306 
            Caption         =   "全屏幕(&F)"
            Shortcut        =   ^X
         End
      End
      Begin VB.Menu a003010 
         Caption         =   "个人化(&I)"
         Shortcut        =   ^Z
      End
      Begin VB.Menu a003011 
         Caption         =   "-"
      End
      Begin VB.Menu a003012 
         Caption         =   "选择 Snowflake(&C)..."
      End
   End
   Begin VB.Menu a004 
      Caption         =   "e"
      Begin VB.Menu s 
         Caption         =   "继续退出时的曲目"
         Shortcut        =   {F5}
      End
      Begin VB.Menu a00402 
         Caption         =   "-"
      End
      Begin VB.Menu ss 
         Caption         =   "查看书签(&B)"
         Enabled         =   0   'False
      End
      Begin VB.Menu dsd 
         Caption         =   "-"
      End
      Begin VB.Menu a00403 
         Caption         =   ""
         Shortcut        =   +{F1}
      End
      Begin VB.Menu a00404 
         Caption         =   ""
         Shortcut        =   +{F2}
      End
      Begin VB.Menu a00405 
         Caption         =   ""
         Shortcut        =   +{F3}
      End
      Begin VB.Menu a004061 
         Caption         =   ""
         Shortcut        =   +{F4}
      End
      Begin VB.Menu a00406 
         Caption         =   ""
         Shortcut        =   +{F5}
      End
      Begin VB.Menu a00407 
         Caption         =   "-"
      End
      Begin VB.Menu a00408 
         Caption         =   "标记书签 [A]"
         Shortcut        =   +^{F1}
      End
      Begin VB.Menu a00409 
         Caption         =   "标记书签 [B]"
         Shortcut        =   +^{F2}
      End
      Begin VB.Menu a004010 
         Caption         =   "标记书签 [C]"
         Shortcut        =   +^{F3}
      End
      Begin VB.Menu a004011 
         Caption         =   "标记书签 [D]"
         Shortcut        =   +^{F4}
      End
      Begin VB.Menu a004012 
         Caption         =   "标记书签 [E]"
         Shortcut        =   +^{F5}
      End
   End
   Begin VB.Menu a005 
      Caption         =   "f"
      Begin VB.Menu a00001 
         Caption         =   "播放所选(&P)"
      End
      Begin VB.Menu rgfdgwqr 
         Caption         =   "-"
      End
      Begin VB.Menu a00002 
         Caption         =   "删除所选项(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu werewr 
         Caption         =   "删除所选文件(&E)"
      End
      Begin VB.Menu a00003 
         Caption         =   "-"
      End
      Begin VB.Menu a00501 
         Caption         =   "添加文件(&A)..."
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu a005021 
         Caption         =   "添加地址(&R)..."
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu a005031 
         Caption         =   "添加文件夹(&F)..."
      End
      Begin VB.Menu a005042 
         Caption         =   "-"
      End
      Begin VB.Menu fewf 
         Caption         =   "查找曲目(&N)..."
         Enabled         =   0   'False
      End
      Begin VB.Menu a005052 
         Caption         =   "整理(&I)"
      End
      Begin VB.Menu dfasdfasdfw 
         Caption         =   "-"
      End
      Begin VB.Menu a005063 
         Caption         =   "导出保存(&S)..."
      End
      Begin VB.Menu kk98 
         Caption         =   "-"
      End
      Begin VB.Menu a005074 
         Caption         =   "清空(&C)"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu a006 
      Caption         =   "g"
      Begin VB.Menu a00601 
         Caption         =   "媒体库(&L)"
         Enabled         =   0   'False
      End
      Begin VB.Menu vdfv 
         Caption         =   "媒体助手(&H)"
         Enabled         =   0   'False
      End
      Begin VB.Menu sdafgerg 
         Caption         =   "-"
      End
      Begin VB.Menu a005010 
         Caption         =   "连接随身听(&M)"
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu a00502 
         Caption         =   "-"
      End
      Begin VB.Menu a00503 
         Caption         =   "从 CD 复制音乐(&C)"
      End
      Begin VB.Menu sdfhgjy 
         Caption         =   "从 VCD、DVD 复制视频(&V)"
      End
      Begin VB.Menu jtrhsgcdgr 
         Caption         =   "-"
      End
      Begin VB.Menu a00505 
         Caption         =   "视频捕获(&A)"
      End
      Begin VB.Menu greyhy 
         Caption         =   "-"
      End
      Begin VB.Menu a00509 
         Caption         =   "刻录到 CD (&B)"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu a007 
      Caption         =   "h"
      Begin VB.Menu a00701 
         Caption         =   "选项(&S)"
         Shortcut        =   ^Q
      End
      Begin VB.Menu a00703 
         Caption         =   "播放选项(&O)..."
      End
      Begin VB.Menu a00702 
         Caption         =   "-"
      End
      Begin VB.Menu fhghg 
         Caption         =   "均衡(&E)"
         Enabled         =   0   'False
      End
      Begin VB.Menu a 
         Caption         =   "图形均衡(&V)"
         Enabled         =   0   'False
      End
      Begin VB.Menu d 
         Caption         =   "音频混合器(&U)"
      End
      Begin VB.Menu a00705 
         Caption         =   "-"
      End
      Begin VB.Menu a00707 
         Caption         =   "许可证(&A)"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu a008 
      Caption         =   "i"
      Begin VB.Menu a00801 
         Caption         =   "帮助(&H)"
         Shortcut        =   ^H
      End
      Begin VB.Menu a00802 
         Caption         =   "-"
      End
      Begin VB.Menu fdf4 
         Caption         =   "电话求助(&T)"
      End
      Begin VB.Menu fewftget 
         Caption         =   "交流和推荐 Snowman Media(&X)"
      End
      Begin VB.Menu dfdsfea5f4as 
         Caption         =   "-"
      End
      Begin VB.Menu a00803 
         Caption         =   "在线更新(&U)"
         Enabled         =   0   'False
      End
      Begin VB.Menu sdfewf 
         Caption         =   "-"
      End
      Begin VB.Menu a00805 
         Caption         =   "访问流动网络(&I)"
         Enabled         =   0   'False
      End
      Begin VB.Menu a00806 
         Caption         =   "电邮联系作者(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu a00807 
         Caption         =   "-"
      End
      Begin VB.Menu a00808 
         Caption         =   "自述(&R)"
      End
      Begin VB.Menu a00809 
         Caption         =   "许可协议(&L)"
      End
      Begin VB.Menu a00810 
         Caption         =   "-"
      End
      Begin VB.Menu a00811 
         Caption         =   "关于 Snomwan Media ilxz(&A)..."
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu b002 
      Caption         =   "j"
      Begin VB.Menu sdf 
         Caption         =   "打开(&O)..."
      End
      Begin VB.Menu dfd 
         Caption         =   "添加到列表(&A)..."
      End
      Begin VB.Menu kjhk 
         Caption         =   "-"
      End
      Begin VB.Menu jkjhl 
         Caption         =   "播放(&P)"
      End
      Begin VB.Menu jlkl 
         Caption         =   "暂停(&U)"
      End
      Begin VB.Menu jhl 
         Caption         =   "停止(&S)"
      End
      Begin VB.Menu jlk 
         Caption         =   "上一首曲目(&B)"
      End
      Begin VB.Menu il 
         Caption         =   "下一首曲目(&F)"
      End
      Begin VB.Menu jhdtg 
         Caption         =   "音量(&V)"
         Begin VB.Menu rthrtb 
            Caption         =   "100% (&9)"
         End
         Begin VB.Menu brt 
            Caption         =   "88% (&8)"
         End
         Begin VB.Menu brthsr 
            Caption         =   "75% (&7)"
         End
         Begin VB.Menu asergrjh 
            Caption         =   "63% (&6)"
         End
         Begin VB.Menu vnrtt 
            Caption         =   "50% (&5)"
         End
         Begin VB.Menu sedgghh 
            Caption         =   "38% (&4)"
         End
         Begin VB.Menu dfhgreff 
            Caption         =   "25% (&3)"
         End
         Begin VB.Menu ertgre 
            Caption         =   "13% (&2)"
         End
         Begin VB.Menu bsrety 
            Caption         =   "0% (&0)"
         End
      End
      Begin VB.Menu dfdfdfdf 
         Caption         =   "-"
      End
      Begin VB.Menu fnbrth 
         Caption         =   "DVD 功能(&D)"
         Enabled         =   0   'False
      End
      Begin VB.Menu cvbtr 
         Caption         =   "-"
      End
      Begin VB.Menu juyjk 
         Caption         =   "Snowflake 模式(&N)"
      End
      Begin VB.Menu vbx 
         Caption         =   "可视效果(&Z)"
         Enabled         =   0   'False
         Begin VB.Menu rtghfg 
            Caption         =   "NEXT"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu erg 
         Caption         =   "全屏幕(&W)"
      End
      Begin VB.Menu ythg 
         Caption         =   "个人化(&L)"
      End
      Begin VB.Menu erth 
         Caption         =   "始终前置(&T)"
      End
      Begin VB.Menu juyu 
         Caption         =   "-"
      End
      Begin VB.Menu kuykuk 
         Caption         =   "统计信息(&I)..."
      End
      Begin VB.Menu juyjuk 
         Caption         =   "文件属性(&R)..."
      End
      Begin VB.Menu juykik 
         Caption         =   "-"
      End
      Begin VB.Menu kuyok 
         Caption         =   "选项(&E)"
         Begin VB.Menu rg 
            Caption         =   "选项(&O)"
         End
         Begin VB.Menu ghjykjyt 
            Caption         =   "播放选项(&P)..."
         End
      End
      Begin VB.Menu uoiooo 
         Caption         =   "帮助(&H)"
         Begin VB.Menu iuot7oi7to 
            Caption         =   "帮助(&H)"
         End
         Begin VB.Menu fsdf 
            Caption         =   "-"
         End
         Begin VB.Menu y5665y 
            Caption         =   "关于 Snowman Media ilxz(&A)..."
         End
      End
      Begin VB.Menu y56yhtdh 
         Caption         =   "-"
      End
      Begin VB.Menu hf6t 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu asdf 
      Caption         =   "k"
      Begin VB.Menu hgj 
         Caption         =   "100% (&9)"
         Shortcut        =   ^{F9}
      End
      Begin VB.Menu rty 
         Caption         =   "88% (&8)"
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu cvb 
         Caption         =   "75% (&7)"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu fghrt 
         Caption         =   "63% (&6)"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu vcbfg 
         Caption         =   "50% (&5)"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu bcvhbfg 
         Caption         =   "38% (&4)"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu bfrg 
         Caption         =   "25% (&3)"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu bfgg 
         Caption         =   "13% (&2)"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu bfgb 
         Caption         =   "0% (&0)"
         Shortcut        =   ^N
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TimeOne As Boolean
Const WM_SETHOTKEY = &H32
'Const HOTKEYF_SHIFT = &H1
Const HOTKEYF_CONTROL = &H2
Const HOTKEYF_ALT = &H4
Dim MoveX As Integer, MoveY As Integer
Dim Info As String
Dim CDRom As String
Dim i As Integer
Dim Pid As Integer
Const LB_ITEMFROMPOINT = &H1A9
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Dim Cd As Boolean
Dim FastForwardSpeed As Long        ' seconds to seek for ff/rew
Dim Playing As Boolean                ' true if CD is currently playing
Dim CDLoad As Boolean                  ' true if CD is the the player
Dim TotalTracks As Integer              ' total tracks tracks on audio CD
Dim TrackLength() As String              ' array containing length of each track
Dim Track As Integer                     ' current track
Dim Minute As Integer                   ' current minute on track
Dim Second As Integer                  ' current second on track
Dim Command As String                 ' string to hold mci command strings
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias _
        "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetSpecialFolderLocation Lib _
        "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder _
        As Long, pIdl As ITEMIDLIST) As Long
Private Declare Function SHGetFileInfo Lib "Shell32" Alias _
        "SHGetFileInfoA" (ByVal pszPath As Any, ByVal _
        dwFileAttributes As Long, psfi As SHFILEINFO, ByVal _
        cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ShellAbout Lib "shell32.dll" Alias _
        "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As _
        String, ByVal szOtherStuff As String, ByVal hIcon As Long) _
        As Long
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
Dim NOERROR As Long
Dim SHGFI_PIDL As Long
Dim SHGFI_ICON As Long
Dim SHGFI_SMALLICON As Long
Dim SearchFlag As Integer    ' Used as flag for cancelling, etc.
Const SPI_SETSCREENSAVEACTIVE = 17
Const SPI_SETSCREENSAVETIMEOUT = 15
Const SPIF_SENDWININICHANGE = &H2
Private Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
     ByVal lpvParam As Long, ByVal fuWinIni As Long) As Long
Dim LsTy As Integer

Private Sub EbSv()
  Call SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, 1, 0, SPIF_SENDWININICHANGE)
End Sub

Private Sub DbSv()
  Call SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, 0, 0, SPIF_SENDWININICHANGE)
End Sub

Private Function GetFolderValue(wIdx As Integer) As Long
    If wIdx < 2 Then
        GetFolderValue = 0
    ElseIf wIdx < 12 Then
        GetFolderValue = wIdx
    Else
        GetFolderValue = wIdx + 4
    End If
End Function

Private Sub a00001_Click()
LF1_DblClick
End Sub

Private Sub a00002_Click()
On Error Resume Next
If LF1.ListCount > 0 And LF1.SelCount > 0 Then
If Pid <= LF1.ListIndex Then Pid = Pid - 1
LF1.RemoveItem (LF1.ListIndex)
End If
End Sub

Private Sub a001010_Click()
Shell App.Path + "\SmM_HDTV.exe", vbNormalFocus
End Sub

Private Sub a00102_Click()
Dim SelectFileName As String
aa:
 SelectFileName = InputBox("请输入万维网地址 (URL) 或指定你要打开的本地文件路径。", , SelectFileName)
  If Len(SelectFileName) > 0 Then
        LF1.Clear
 If Ly.FileExists(SelectFileName) = True Then
      LF1.AddItem SelectFileName, 0
      Pid = 0
      LF1.ListIndex = Pid
      LF1_DblClick
 Else
  If MsgBox("所请求的媒体文件不存在,如果是万维网资源请先连接网络。", vbRetryCancel) = vbRetry Then
     GoTo aa:
   Else
   Exit Sub
   End If
   End If
     End If
End Sub

Private Sub a00103_Click()
  On Error Resume Next
  Dim BI As BROWSEINFO
  Dim nFolder As Long
  Dim IDL As ITEMIDLIST
  Dim pIdl As Long
  Dim sPath As String
  Dim SHFI As SHFILEINFO
  Dim m_wCurOptIdx As Integer
  Dim txtPath As String
  Dim txtDisplayName As String
    With BI
    .hOwner = Me.hWnd
    nFolder = GetFolderValue(m_wCurOptIdx)
     If SHGetSpecialFolderLocation(ByVal Me.hWnd, ByVal nFolder, IDL) = NOERROR Then
      .pidlRoot = IDL.mkid.cb
    End If
     .pszDisplayName = String$(MAX_PATH, 0)
    .lpszTitle = "请选择你要打开的媒体文件夹。"
    .ulFlags = 0
  End With
  txtPath = ""
  txtDisplayName = ""
    pIdl = SHBrowseForFolder(BI)
  If pIdl = 0 Then Exit Sub
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
   If Len(txtPath) > 0 Then
                Me.MousePointer = 11
     If Dir1.ListCount <> 0 Or File1.ListCount <> 0 Then
                   LF1.Clear
                 File1.Path = txtPath
  If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AllFiles") = True Then
            Dim result As Integer
Dim firstpath As String, dircount As Integer
Dir1.Path = txtPath
firstpath = Dir1.Path
dircount = Dir1.ListCount
result = DirDiver(firstpath, dircount, "")
Pid = 0
      LF1.ListIndex = Pid
     If LF1.ListCount > 0 Then
     LF1_DblClick
    Else
         MsgBox "文件夹内无可用媒体,请重新指定路径。", vbExclamation
          End If
   Else
   For i = 0 To File1.ListCount - 1
     LF1.AddItem File1.Path + "\" + File1.List(i), i
    Next
      Pid = 0
      LF1.ListIndex = Pid
     If LF1.ListCount > 0 Then
     LF1_DblClick
    Else
         MsgBox "文件夹内无可用媒体,请重新指定路径。", vbExclamation
          End If
     End If
    End If
 Me.MousePointer = 0
   End If
End Sub

Private Sub a00105_Click()
If Ly.FileExists(CDRom + "Track01.cda") = True Then
 LF1.Clear
   File1.Path = CDRom
          For i = 0 To File1.ListCount - 1
      LF1.AddItem File1.Path + File1.List(i), i
            Next
 If LF1.ListCount > 0 Then
      Pid = 0
      LF1.ListIndex = Pid
     LF1_DblClick
   End If

Else


MsgBox "找不到音频 CD 唱片,请重新插入。", vbExclamation
End If
End Sub

Private Sub a00106_Click()
On Error Resume Next
If Ly.FileExists(CDRom + "MPEGAV\AVSEQ01.DAT") = True Or Ly.FileExists(CDRom + "MPEGAV\MUSIC01.DAT") = True Then
     LF1.Clear
File1.Path = CDRom + "MPEGAV"
 For i = 0 To File1.ListCount - 1
      LF1.AddItem File1.Path + "\" + File1.List(i), i
Next
 If LF1.ListCount > 0 Then
      Pid = 0
      LF1.ListIndex = Pid
     LF1_DblClick
   End If
Else
 MsgBox "找不到 VCD 视频光盘,请重新插入。", vbExclamation
End If
End Sub


Private Sub a00107_Click()
Shell (App.Path + "\SmM_DVD.exe"), vbNormalFocus
End Sub

Private Sub a00201_Click()
If MP1.ShowCaptioning = False Then
MP1.ShowCaptioning = True
a00201.Checked = True
Ig8.Visible = False
Form_Resize
Else
 MP1.ShowCaptioning = False
a00201.Checked = False
Ig8.Visible = True
End If
End Sub

Private Sub a00202_Click()
If Len(MP1.Filename) > 0 Then Ly.ShowProp MP1.Filename, Me
If Playing = True Then
 If Len(Str(Track)) = 2 Then Ly.ShowProp CDRom + "Track0" + Right(Str(Track), 1) + ".cda", Me
 If Len(Str(Track)) = 3 Then Ly.ShowProp CDRom + "Track" + Right(Str(Track), 2) + ".cda", Me
End If
End Sub

Private Sub a002020_Click()
MP1.ShowDialog mpShowDialogStatistics
End Sub

Private Sub a00301_Click()
Me.Height = 5145
Me.Width = 7545
End Sub

Private Sub a003010_Click()
If a003010.Checked = False Then
Form100.Show
a003010.Checked = True
ythg.Checked = True
Else
Unload Form100
a003010.Checked = False
ythg.Checked = False
End If
End Sub

Private Sub a003012_Click()
Form3.Show
End Sub

Private Sub a00302_Click()
Me.Width = 15090
Me.Height = 10290
End Sub

Private Sub a00303_Click()
Me.WindowState = 2
End Sub

Private Sub a00304_Click()
Me.WindowState = 1
End Sub

Private Sub a00308_Click()
If a00308.Checked = False Then
a00308.Checked = True
erth.Checked = True
Ly.MakeTop Me, True
Else
a00308.Checked = False
erth.Checked = False
Ly.MakeTop Me, False
End If
End Sub

Private Sub a00306_Click()
MP1.DisplaySize = mpFullScreen
mjjhjm.Checked = False
cerde.Checked = False
cbcxbf.Checked = False
a00306.Checked = True
xfg.Checked = False
erg.Checked = True
End Sub

Public Sub RushBm()
On Error Resume Next
If Ly.FileExists(myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_X", "")) = True Then
s.Enabled = True
s.Caption = "最后位置 : " + myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_X_I", "")
Else: s.Enabled = False
s.Caption = "继续退出时的曲目"
End If
If Ly.FileExists(myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_C", "")) = True Then
a00405.Enabled = True
a00405.Caption = "书签 [C] : " + myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_C_I", "")
Else: a00405.Enabled = False
a00405.Caption = "书签 [C] : 无内容"
End If
If Ly.FileExists(myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_D", "")) = True Then
a004061.Enabled = True
a004061.Caption = "书签 [D] : " + myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_D_I", "")
Else: a004061.Enabled = False
a004061.Caption = "书签 [D] : 无内容"
End If
If Ly.FileExists(myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_B", "")) = True Then
a00404.Enabled = True
a00404.Caption = "书签 [B] : " + myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_B_I", "")
Else: a00404.Enabled = False
a00404.Caption = "书签 [B] : 无内容"
End If
If Ly.FileExists(myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_A", "")) = True Then
a00403.Enabled = True
a00403.Caption = "书签 [A] : " + myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_A_I", "")
Else: a00403.Enabled = False
a00403.Caption = "书签 [A] : 无内容"
End If
If Ly.FileExists(myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_E", "")) = True Then
a00406.Enabled = True
a00406.Caption = "书签 [E] : " + myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_E_I", "")
Else: a00406.Enabled = False
a00406.Caption = "书签 [E] : 无内容"
End If
End Sub

Private Sub a004010_Click()
On Error Resume Next
Dim text As String
If Len(MP1.Filename) > 0 Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_C", MP1.CurrentPosition
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_C", MP1.GetMediaInfoString(mpClipFilename)
text = ""
If Len(MP1.GetMediaInfoString(mpClipAuthor)) > 0 Then
text = MP1.GetMediaInfoString(mpClipAuthor)
text = text + " - "
End If
If Len(MP1.GetMediaInfoString(mpClipTitle)) > 0 Then text = text + MP1.GetMediaInfoString(mpClipTitle)
If Len(MP1.GetMediaInfoString(mpClipTitle)) = 0 And Len(MP1.GetMediaInfoString(mpClipTitle)) = 0 Then text = MP1.GetMediaInfoString(mpClipFilename)
text = text + " -" + Str(CLng(MP1.CurrentPosition)) + " 秒 "
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_C_I", text
RushBm
End If
End Sub

Private Sub a004011_Click()
On Error Resume Next
Dim text As String
If Len(MP1.Filename) > 0 Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_D", MP1.CurrentPosition
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_D", MP1.GetMediaInfoString(mpClipFilename)
text = ""
If Len(MP1.GetMediaInfoString(mpClipAuthor)) > 0 Then
text = MP1.GetMediaInfoString(mpClipAuthor)
text = text + " - "
End If
If Len(MP1.GetMediaInfoString(mpClipTitle)) > 0 Then text = text + MP1.GetMediaInfoString(mpClipTitle)
If Len(MP1.GetMediaInfoString(mpClipTitle)) = 0 And Len(MP1.GetMediaInfoString(mpClipTitle)) = 0 Then text = MP1.GetMediaInfoString(mpClipFilename)
text = text + " -" + Str(CLng(MP1.CurrentPosition)) + " 秒 "
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_D_I", text
RushBm
End If
End Sub

Private Sub a004012_Click()
On Error Resume Next
Dim text As String
If Len(MP1.Filename) > 0 Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_E", MP1.CurrentPosition
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_E", MP1.GetMediaInfoString(mpClipFilename)
text = ""
If Len(MP1.GetMediaInfoString(mpClipAuthor)) > 0 Then
text = MP1.GetMediaInfoString(mpClipAuthor)
text = text + " - "
End If
If Len(MP1.GetMediaInfoString(mpClipTitle)) > 0 Then text = text + MP1.GetMediaInfoString(mpClipTitle)
If Len(MP1.GetMediaInfoString(mpClipTitle)) = 0 And Len(MP1.GetMediaInfoString(mpClipTitle)) = 0 Then text = MP1.GetMediaInfoString(mpClipFilename)
text = text + " -" + Str(CLng(MP1.CurrentPosition)) + " 秒 "
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_E_I", text
RushBm
End If
End Sub

Private Sub a00403_Click()
   If Ly.FileExists(myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_A", "")) = True Then
     MP1.Filename = myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_A", "")
     MP1.CurrentPosition = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_A")
       LF1.Clear
     LF1.AddItem MP1.Filename, 0
       Else
     MsgBox "找不到媒体文件,文件可能已被移动或改名。", vbExclamation
     End If
End Sub

Private Sub a00404_Click()
  If Ly.FileExists(myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_B", "")) = True Then
     MP1.Filename = myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_B", "")
     MP1.CurrentPosition = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_B")
          LF1.Clear
     LF1.AddItem MP1.Filename, 0
    Else
     MsgBox "找不到媒体文件,文件可能已被移动或改名。", vbExclamation
     End If
End Sub

Private Sub a00405_Click()
  If Ly.FileExists(myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_C", "")) = True Then
     MP1.Filename = myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_C", "")
     MP1.CurrentPosition = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_C")
          LF1.Clear
     LF1.AddItem MP1.Filename, 0
    Else
    MsgBox "找不到媒体文件,文件可能已被移动或改名。", vbExclamation
   End If
End Sub

Private Sub a00406_Click()
   If Ly.FileExists(myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_E", "")) = True Then
     MP1.Filename = myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_E", "")
     MP1.CurrentPosition = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_E")
          LF1.Clear
     LF1.AddItem MP1.Filename, 0
    Else
     MsgBox "找不到媒体文件,文件可能已被移动或改名。", vbExclamation
    End If
End Sub

Private Sub a004061_Click()
   If Ly.FileExists(myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_D", "")) = True Then
     MP1.Filename = myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_D", "")
     MP1.CurrentPosition = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_D")
       LF1.Clear
     LF1.AddItem MP1.Filename, 0
     Else
     MsgBox "找不到媒体文件,文件可能已被移动或改名。", vbExclamation
End If
End Sub

Private Sub a00408_Click()
On Error Resume Next
Dim text As String
If Len(MP1.Filename) > 0 Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_A", MP1.CurrentPosition
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_A", MP1.GetMediaInfoString(mpClipFilename)
text = ""
If Len(MP1.GetMediaInfoString(mpClipAuthor)) > 0 Then
text = MP1.GetMediaInfoString(mpClipAuthor)
text = text + " - "
End If
If Len(MP1.GetMediaInfoString(mpClipTitle)) > 0 Then text = text + MP1.GetMediaInfoString(mpClipTitle)
If Len(MP1.GetMediaInfoString(mpClipTitle)) = 0 And Len(MP1.GetMediaInfoString(mpClipTitle)) = 0 Then text = MP1.GetMediaInfoString(mpClipFilename)
text = text + " -" + Str(CLng(MP1.CurrentPosition)) + " 秒 "
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_A_I", text
RushBm
End If
End Sub

Private Sub a00409_Click()
On Error Resume Next
Dim text As String
If Len(MP1.Filename) > 0 Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_B", MP1.CurrentPosition
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_B", MP1.GetMediaInfoString(mpClipFilename)
text = ""
If Len(MP1.GetMediaInfoString(mpClipAuthor)) > 0 Then
text = MP1.GetMediaInfoString(mpClipAuthor)
text = text + " - "
End If
If Len(MP1.GetMediaInfoString(mpClipTitle)) > 0 Then text = text + MP1.GetMediaInfoString(mpClipTitle)
If Len(MP1.GetMediaInfoString(mpClipTitle)) = 0 And Len(MP1.GetMediaInfoString(mpClipTitle)) = 0 Then text = MP1.GetMediaInfoString(mpClipFilename)
text = text + " -" + Str(CLng(MP1.CurrentPosition)) + " 秒 "
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_B_I", text
RushBm
End If
End Sub

Sub SetBookMark()
On Error Resume Next
Dim text As String
If Len(MP1.Filename) > 0 Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_X", MP1.CurrentPosition
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_X", MP1.GetMediaInfoString(mpClipFilename)
text = ""
If Len(MP1.GetMediaInfoString(mpClipAuthor)) > 0 Then
text = MP1.GetMediaInfoString(mpClipAuthor)
text = text + " - "
End If
If Len(MP1.GetMediaInfoString(mpClipTitle)) > 0 Then text = text + MP1.GetMediaInfoString(mpClipTitle)
If Len(MP1.GetMediaInfoString(mpClipTitle)) = 0 And Len(MP1.GetMediaInfoString(mpClipTitle)) = 0 Then text = MP1.GetMediaInfoString(mpClipFilename)
text = text + " -" + Str(CLng(MP1.CurrentPosition)) + " 秒 "
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_X_I", text
RushBm
End If
End Sub

Private Sub a005021_Click()
Dim SelectFileName As String
aa:
 SelectFileName = InputBox("请输入万维网地址 (URL) 或指定你要打开的本地文件路径。", , SelectFileName)
  If Len(SelectFileName) > 0 Then
      If Ly.FileExists(SelectFileName) = True Then
         LF1.AddItem SelectFileName, LF1.ListCount
         a005052_Click
      Else
      If MsgBox("所请求的资源文件不存在,如果是万维网资源请先连接网络。", vbRetryCancel) = vbRetry Then
     GoTo aa:
             Else
                 Exit Sub
       End If
       End If
End If
End Sub

Private Sub a00503_Click()
If Playing = True Then Image6_MouseUp 1, 0, 0, 0
Shell (App.Path + "\SmM_Casket\SmM_Casket.exe"), vbNormalFocus
End Sub

Private Sub a005031_Click()
    Dim BI As BROWSEINFO
  Dim nFolder As Long
  Dim IDL As ITEMIDLIST
  Dim pIdl As Long
  Dim sPath As String
  Dim SHFI As SHFILEINFO
  Dim m_wCurOptIdx As Integer
  Dim txtPath As String
  Dim txtDisplayName As String
    With BI
    .hOwner = Me.hWnd
    nFolder = GetFolderValue(m_wCurOptIdx)
     If SHGetSpecialFolderLocation(ByVal Me.hWnd, ByVal nFolder, IDL) = NOERROR Then
      .pidlRoot = IDL.mkid.cb
    End If
     .pszDisplayName = String$(MAX_PATH, 0)
    .lpszTitle = "请选择你要添加的媒体文件夹。"
    .ulFlags = 0
  End With
   txtPath = ""
  txtDisplayName = ""
  pIdl = SHBrowseForFolder(BI)
  If pIdl = 0 Then Exit Sub
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
   If Len(txtPath) > 0 Then
                Me.MousePointer = 11
                 File1.Path = txtPath
     If Dir1.ListCount <> 0 Or File1.ListCount <> 0 Then
     If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AllFiles") = True Then
            Dim result As Integer
Dim firstpath As String, dircount As Integer
Dir1.Path = txtPath
firstpath = Dir1.Path
dircount = Dir1.ListCount
result = DirDiver(firstpath, dircount, "")
   Else
   For i = 0 To File1.ListCount - 1
     LF1.AddItem File1.Path + "\" + File1.List(i), LF1.ListCount
    Next
     End If
    End If
 Me.MousePointer = 0
   End If
End Sub

Private Sub a005052_Click()
On Error Resume Next
Me.MousePointer = 11
Dim text As String
Dim Tid As Integer
Tid = 0
For Tid = 0 To LF1.ListCount - 1
  If Right(LF1.List(Tid), 4) = ".sml" Or Right(LF1.List(Tid), 4) = ".SML" Then
  Open LF1.List(Tid) For Input As #1
    While Not EOF(1)
    Line Input #1, text
    LF1.AddItem RTrim(text)
    Wend
    Close #1
  LF1.RemoveItem (Tid)
  Tid = Tid - 1
  End If
 Next
For Tid = 0 To LF1.ListCount - 1
    If Ly.FileExists(LF1.List(Tid)) = False Then
       For i = Tid + 1 To LF1.ListCount - 1
          If LF1.List(i) = LF1.List(Tid) Then
          LF1.RemoveItem (i)
          i = i - 1
          End If
        Next i
      LF1.RemoveItem (Tid)
    Tid = Tid - 1
    End If
Next Tid
For Tid = 0 To LF1.ListCount - 1
    For i = Tid + 1 To LF1.ListCount - 1
        If LF1.List(i) = LF1.List(Tid) Then
         LF1.RemoveItem (i)
        i = i - 1
        End If
      Next i
Next Tid
If Pid > LF1.ListCount - 1 Then Pid = 0
Me.MousePointer = 0
End Sub

Private Sub a005063_Click()
If LF1.ListCount > 0 Then
      CommonDialog1.Filename = ""
        CommonDialog1.DialogTitle = "输入你要保存的文件名"
          CommonDialog1.Filter = "列表文件 (*.sml)" & _
          "|*.sml|所有文件 (*.*)|*.*"
             CommonDialog1.ShowSave
          If Len(CommonDialog1.Filename) > 0 Then
            If Right(CommonDialog1.Filename, 4) <> ".sml" Then CommonDialog1.Filename = CommonDialog1.Filename + ".sml"
               If Ly.FileExists(CommonDialog1.Filename) = True Then
                      If MsgBox("文件 """ + CommonDialog1.Filename + """ 已存在,要覆盖吗?", vbYesNo) = vbYes Then
                                   Open CommonDialog1.Filename For Output As #1
                                    For i = 0 To LF1.ListCount - 1
                                    Print #1, LF1.List(i)
                                     Next i
                                    Close (1)
                               Else
                               Exit Sub
                               End If
                        Else
                            Open CommonDialog1.Filename For Output As #1
                                    For i = 0 To LF1.ListCount - 1
                                    Print #1, LF1.List(i)
                                     Next i
                                    Close (1)
                             End If
                   End If
        End If
End Sub

Private Sub a00701_Click()
Shell (App.Path + "\SmM_Settings.exe"), vbNormalFocus
End Sub

Private Sub a00703_Click()
MP1.ShowDialog mpShowDialogOptions
End Sub

Private Sub a00801_Click()
Shell (App.Path + "\SmM_Help.exe")
End Sub

Private Sub a00806_Click()
Ly.SendMail ("leask@21cn.com")
End Sub

Private Sub a00808_Click()
Ly.OpenFile (App.Path + "\自述文件.txt")
End Sub

Private Sub a00809_Click()
Ly.OpenFile (App.Path + "\许可协议.txt")
End Sub

Private Sub a00811_Click()
Form2.Show
End Sub

Private Sub asergrjh_Click()
fghrt_Click
End Sub

Private Sub bcvhbfg_Click()
HB.SetVolume 21, 21, 0
bfgb.Checked = False
 bfgg.Checked = False
 bfrg.Checked = False
 bcvhbfg.Checked = True
  vcbfg.Checked = False
 fghrt.Checked = False
 cvb.Checked = False
  rty.Checked = False
  hgj.Checked = False
    bsrety.Checked = False
 ertgre.Checked = False
 dfhgreff.Checked = False
  sedgghh.Checked = True
 vnrtt.Checked = False
  asergrjh.Checked = False
 brthsr.Checked = False
 brt.Checked = False
 rthrtb.Checked = False
End Sub

Private Sub bfgb_Click()
HB.SetVolume 0, 0, 0
bfgb.Checked = True
 bfgg.Checked = False
 bfrg.Checked = False
 bcvhbfg.Checked = False
  vcbfg.Checked = False
 fghrt.Checked = False
 cvb.Checked = False
  rty.Checked = False
  hgj.Checked = False
  bsrety.Checked = True
 ertgre.Checked = False
 dfhgreff.Checked = False
  sedgghh.Checked = False
 vnrtt.Checked = False
  asergrjh.Checked = False
 brthsr.Checked = False
 brt.Checked = False
 rthrtb.Checked = False
End Sub

Private Sub bfgg_Click()
HB.SetVolume 7, 7, 0
bfgb.Checked = False
 bfgg.Checked = True
 bfrg.Checked = False
 bcvhbfg.Checked = False
  vcbfg.Checked = False
 fghrt.Checked = False
 cvb.Checked = False
  rty.Checked = False
  hgj.Checked = False
  bsrety.Checked = False
 ertgre.Checked = True
 dfhgreff.Checked = False
  sedgghh.Checked = False
 vnrtt.Checked = False
  asergrjh.Checked = False
 brthsr.Checked = False
 brt.Checked = False
 rthrtb.Checked = False
End Sub

Private Sub bfrg_Click()
HB.SetVolume 14, 14, 0
bfgb.Checked = False
 bfgg.Checked = False
 bfrg.Checked = True
 bcvhbfg.Checked = False
  vcbfg.Checked = False
 fghrt.Checked = False
 cvb.Checked = False
  rty.Checked = False
  hgj.Checked = False
  bsrety.Checked = False
 ertgre.Checked = False
 dfhgreff.Checked = True
  sedgghh.Checked = False
 vnrtt.Checked = False
  asergrjh.Checked = False
 brthsr.Checked = False
 brt.Checked = False
 rthrtb.Checked = False
End Sub

Private Sub brt_Click()
rty_Click
End Sub

Private Sub brthsr_Click()
cvb_Click
End Sub

Private Sub bsrety_Click()
 bfgb_Click
End Sub

Private Sub cbcxbf_Click()
MP1.DisplaySize = mpDefaultSize
mjjhjm.Checked = False
cerde.Checked = False
cbcxbf.Checked = True
a00306.Checked = False
xfg.Checked = False
erg.Checked = False
End Sub

Private Sub cerde_Click()
MP1.DisplaySize = mpHalfSize
mjjhjm.Checked = False
cerde.Checked = True
cbcxbf.Checked = False
a00306.Checked = False
xfg.Checked = False
erg.Checked = False
End Sub


Private Sub Cl0_GotFocus()
cLT1.SetFocus
End Sub

Private Sub cLT1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub cLT1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Form1.Left = Form1.Left + (X - MoveX)
Form1.Top = Form1.Top + (Y - MoveY)
End Sub

Private Sub cLT1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X + cLT1.Left, Y + cLT1.Top
End Sub













Private Sub cvb_Click()
HB.SetVolume 42, 42, 0
bfgb.Checked = False
 bfgg.Checked = False
 bfrg.Checked = False
 bcvhbfg.Checked = False
  vcbfg.Checked = False
 fghrt.Checked = False
 cvb.Checked = True
  rty.Checked = False
  hgj.Checked = False
  bsrety.Checked = False
 ertgre.Checked = False
 dfhgreff.Checked = False
  sedgghh.Checked = False
 vnrtt.Checked = False
  asergrjh.Checked = False
 brthsr.Checked = True
 brt.Checked = False
 rthrtb.Checked = False
End Sub

Private Sub d_Click()
If Ly.FileExists(Ly.GetSysPath + "\sndvol32.exe") = True Then Shell (Ly.GetSysPath + "\sndvol32.exe"), vbNormalFocus
If Ly.FileExists(Ly.GetWinPath + "\sndvol32.exe") = True Then Shell (Ly.GetWinPath + "\sndvol32.exe"), vbNormalFocus
End Sub

Private Sub dfd_Click()
a00501_Click
End Sub

Private Sub dfg_Click()
Image10_MouseUp 1, 0, 0, 0
End Sub

Private Sub dfgr_Click()
Image7_MouseUp 1, 0, 0, 0
End Sub

Private Sub dfhgreff_Click()
bfrg_Click
End Sub

Private Sub dfhgrtht_Click()
Image5_MouseUp 1, 0, 0, 0
End Sub

Private Sub Dir1_Change()
    ' Update File listbox to sync with Dir listbox.
    File1.Path = Dir1.Path
End Sub

Private Sub a00108_Click()
Dim result As Integer
Me.MousePointer = 11
LF1.Clear
Dim firstpath As String, dircount As Integer
Dir1.Path = CDRom
firstpath = Dir1.Path
dircount = Dir1.ListCount
result = DirDiver(firstpath, dircount, "")
Me.MousePointer = 0
Pid = 0
      LF1.ListIndex = Pid
     LF1_DblClick
End Sub

Private Sub a002045_Click()
If a002045.Checked = False Then
 a002045.Checked = True
 Ig12.Picture = LoadPicture(App.Path + "\SmM_Icos\11.gif")
 Ig12.ToolTipText = "随机"
Else
  a002045.Checked = False
   Ig12.Picture = LoadPicture(App.Path + "\SmM_Icos\10.gif")
  Ig12.ToolTipText = "原序"
End If
End Sub

Private Sub a00205_Click()
If a00205.Checked = False Then
 a00205.Checked = True
 Ig13.Picture = LoadPicture(App.Path + "\SmM_Icos\03.gif")
 Ig13.ToolTipText = "循环"
Else
  a00205.Checked = False
   Ig13.Picture = LoadPicture(App.Path + "\SmM_Icos\02.gif")
 Ig13.ToolTipText = "不循环"
End If
End Sub

Private Sub a00501_Click()
CommonDialog1.Filename = ""
          CommonDialog1.ShowOpen
        If Len(CommonDialog1.Filename) > 0 Then
      LF1.AddItem CommonDialog1.Filename, LF1.ListCount
       a005052_Click
 End If
 End Sub

Private Sub a00101_Click()
CommonDialog1.Filename = ""
CommonDialog1.DialogTitle = "浏览要播放的媒体"
          CommonDialog1.ShowOpen
          If Len(CommonDialog1.Filename) > 0 Then
        LF1.Clear
      LF1.AddItem CommonDialog1.Filename, 0
      Pid = 0
      LF1.ListIndex = Pid
      LF1_DblClick
     End If
 End Sub

Private Sub a00505_Click()
Shell (App.Path + "\SmM_Capturer.exe"), vbNormalFocus
End Sub

Private Sub a005074_Click()
LF1.Clear
End Sub

Private Sub erg_Click()
If erg.Checked = False Then
a00306_Click
Else
mjjhjm_Click
End If
End Sub

Private Sub ertgre_Click()
bfgg_Click
End Sub

Private Sub erth_Click()
a00308_Click
End Sub

Private Sub fdf4_Click()
Dim SelectFileName As String
 SelectFileName = InputBox("请输入你要拨打的求助电话。", , SelectFileName)
  If Len(SelectFileName) > 0 Then
   Ly.Dial SelectFileName, Me
 Else
   Exit Sub
End If
End Sub

Private Sub fewftget_Click()
Shell (App.Path + "\SmM_Tell.exe"), vbNormalFocus
End Sub

Sub SetVw()
If cerde.Checked = True Then mjjhjm_Click
If MP1.ImageSourceWidth = 0 Then Exit Sub
Me.Width = MP1.ImageSourceWidth * 15 + 3150
Me.Height = MP1.ImageSourceHeight * 15 + 2050
End Sub
Private Sub fghgt_Click()
If fghgt.Checked = False Then
   fghgt.Checked = True
   SetVw
Else
   fghgt.Checked = False
End If
End Sub

Private Sub fghrt_Click()
HB.SetVolume 35, 35, 0
bfgb.Checked = False
 bfgg.Checked = False
 bfrg.Checked = False
 bcvhbfg.Checked = False
  vcbfg.Checked = False
 fghrt.Checked = True
 cvb.Checked = False
  rty.Checked = False
  hgj.Checked = False
  bsrety.Checked = False
 ertgre.Checked = False
 dfhgreff.Checked = False
  sedgghh.Checked = False
 vnrtt.Checked = False
  asergrjh.Checked = True
 brthsr.Checked = False
 brt.Checked = False
 rthrtb.Checked = False
 End Sub

Private Sub fgrgfrgrgrtg_Click()
Image8_MouseUp 1, 0, 0, 0
End Sub

Private Sub Fm0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub Fm0_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Form1.Left = Form1.Left + (X - MoveX)
Form1.Top = Form1.Top + (Y - MoveY)
End Sub

Private Sub Fm0_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X + Fm0.Left, Y + Fm0.Top
End Sub

Private Sub Fm0_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Fm10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub Fm10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Form1.Left = Form1.Left + (X - MoveX)
Form1.Top = Form1.Top + (Y - MoveY)
End Sub

Private Sub Fm10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X, Y + Fm10.Top
End Sub

Private Sub Fm10_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Fm5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub Fm5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Form1.Left = Form1.Left + (X - MoveX)
Form1.Top = Form1.Top + (Y - MoveY)
End Sub

Private Sub Fm5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X + Fm5.Left, Y + Fm5.Top
End Sub

Private Sub Fm5_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X, Y
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Form1.Left = Form1.Left + (X - MoveX)
Form1.Top = Form1.Top + (Y - MoveY)
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, 90 + X, Frame1.Top + Y
End Sub

Private Sub Frame1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub ggg_Click()
Dim SwfFile As String
SwfFile = CommonDialog1.Filter
CommonDialog1.Filter = "Flash 影片 (*.swf)|*.swf|所有文件 (*.*)|*.*"
CommonDialog1.Filename = ""
CommonDialog1.DialogTitle = "浏览要播放的 Flash 影片"
          CommonDialog1.ShowOpen
          If Len(CommonDialog1.Filename) > 0 Then
 Shell (App.Path + "\SmM_Flash.exe " + CommonDialog1.Filename), vbNormalFocus
     End If
CommonDialog1.Filter = SwfFile
End Sub

Private Sub ghjykjyt_Click()
a00703_Click
End Sub

Private Sub grdfg_Click()
Image4_MouseUp 1, 0, 0, 0
End Sub

Private Sub greg_Click()
Image3_MouseUp 1, 0, 0, 0
End Sub

Private Sub hf6t_Click()
Unload Me
End Sub

Private Sub hgj_Click()
HB.SetVolume 53, 53, 0
bfgb.Checked = False
 bfgg.Checked = False
 bfrg.Checked = False
 bcvhbfg.Checked = False
  vcbfg.Checked = False
 fghrt.Checked = False
 cvb.Checked = False
  rty.Checked = False
  hgj.Checked = True
  bsrety.Checked = False
 ertgre.Checked = False
 dfhgreff.Checked = False
  sedgghh.Checked = False
 vnrtt.Checked = False
  asergrjh.Checked = False
 brthsr.Checked = False
 brt.Checked = False
 rthrtb.Checked = True
End Sub

Private Sub Ig0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub Ig0_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Form1.Left = Form1.Left + (X - MoveX)
Form1.Top = Form1.Top + (Y - MoveY)
End Sub

Private Sub Ig0_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X, Y
End Sub


Private Sub Ig0_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Ig1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub Ig1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Form1.Left = Form1.Left + (X - MoveX)
Form1.Top = Form1.Top + (Y - MoveY)
End Sub

Private Sub Ig1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X, Y + Ig1.Top
End Sub

Private Sub Ig1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Ig11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim Info2 As String
If Button = 1 Then
If Len(MP1.Filename) = 0 And Playing = False Then
sdtg56_Click
Exit Sub
End If
If Len(MP1.Filename) > 0 Or Playing = True Then
Info2 = ""
If Len(MP1.Filename) > 0 Then
Info2 = "列表   :" + Str(Pid + 1) + " -" + Str(LF1.ListCount) + vbCrLf
If Len(MP1.GetMediaInfoString(mpClipTitle)) > 0 Then Info2 = Info2 + "标题   : " + MP1.GetMediaInfoString(mpClipTitle) + vbCrLf
If Len(MP1.GetMediaInfoString(mpClipAuthor)) > 0 Then Info2 = Info2 + "艺术家 : " + MP1.GetMediaInfoString(mpClipAuthor) + vbCrLf
If Len(MP1.GetMediaInfoString(mpClipCopyright)) > 0 Then Info2 = Info2 + "版权   : " + MP1.GetMediaInfoString(mpClipCopyright) + vbCrLf
If Len(MP1.GetMediaInfoString(mpClipDescription)) > 0 Then Info2 = Info2 + "描述   : " + MP1.GetMediaInfoString(mpClipDescription) + vbCrLf
If MP1.Bandwidth > 0 Then Info2 = Info2 + "质量   :" + Str(Int(MP1.Bandwidth / 1000)) + " 千字节每秒" + vbCrLf
If MP1.ImageSourceWidth > 0 Then
    Info2 = Info2 + "视频   :" + Str(MP1.ImageSourceWidth) + " x" + Str(MP1.ImageSourceHeight) + " @" + Ly.GetDisplay + vbCrLf
Else
    Info2 = Info2 + "类型   : 仅含音频" + vbCrLf
End If
Info2 = Info2 + "地址   : " + MP1.GetMediaInfoString(mpClipFilename)
End If
If Playing = True Then
 Info2 = "音频 CD @ 驱动器:" + Left(CDRom, 1) + vbCrLf + "曲目   : " + Str(Track) + "/" + TrackTime.Caption + " - " + TotalTrack.Caption + vbCrLf + "标题   : 未知标题" + vbCrLf + "艺术家 : 未知艺术家" + vbCrLf + "唱片集 : 未知唱片集" + vbCrLf + "流派   : 未知流派" + vbCrLf + "标识   : " + Ly.GetDiskNumber(CDRom)
End If
MsgBox Info2, vbInformation
End If
End If
If Button = 2 Then PopupMenu Me.b002, 0, X + Ig11.Left, Y + Ig11.Top
End Sub

Private Sub Ig11_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Ig12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a002045_Click
End If
If Button = 2 Then PopupMenu Me.b002, 0, X + Ig12.Left, Y + Ig12.Top
End Sub

Private Sub Ig12_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub


Private Sub Ig13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a00205_Click
End If
If Button = 2 Then PopupMenu Me.b002, 0, X + Ig13.Left, Y + Ig13.Top
End Sub

Private Sub Ig13_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Ig2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub Ig2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Form1.Left = Form1.Left + (X - MoveX)
Form1.Top = Form1.Top + (Y - MoveY)
End Sub

Private Sub Ig2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X, Y + Ig2.Top
End Sub

Private Sub Ig2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Ig4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub Ig4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Form1.Left = Form1.Left + (X - MoveX)
Form1.Top = Form1.Top + (Y - MoveY)
End Sub

Private Sub Ig4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X + 45, Y + Fm10.Top
End Sub

Private Sub Ig4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Ig7_DblClick()
a00306_Click
End Sub

Private Sub Ig7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub Ig7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Form1.Left = Form1.Left + (X - MoveX)
Form1.Top = Form1.Top + (Y - MoveY)
End Sub

Private Sub Ig7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X + Ig7.Left + Fm0.Left, Y + Ig7.Top + Fm0.Top
End Sub

Private Sub Ig7_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Ig8_DblClick()
a00306_Click
End Sub

Private Sub Ig8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub Ig8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Form1.Left = Form1.Left + (X - MoveX)
Form1.Top = Form1.Top + (Y - MoveY)
End Sub

Private Sub Ig8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X + Ig8.Left + Fm0.Left, Y + Ig8.Top + Fm0.Top
End Sub

Private Sub Ig8_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub il_Click()
Image8_MouseUp 1, 0, 0, 0
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Form1.Left = Form1.Left + (X - MoveX)
Form1.Top = Form1.Top + (Y - MoveY)
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X + Fm5.Left, Y + Fm5.Top
End Sub

Private Sub Image1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub


Private Sub Image10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If Left(MP1.Filename, 3) = "G:\" Or Left(MP1.Filename, 3) = "g:\" Then MP1.Filename = "ilxz"
SendMCIString "set cd door open", True
End If
If Button = 2 Then PopupMenu Me.b002, 0, X + 2025, Y + Fm10.Top
End Sub

Private Sub Image10_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If Cd = True Then
Dim from As String
          If (Track > 1) Then
             from = CStr(Track - 1)
                     Else
      from = CStr(TotalTracks)
         End If
     If (Playing) Then
       Command = "play cd from " & from
       SendMCIString Command, True
      Else
       Command = "seek cd to " & from
       SendMCIString Command, True
    End If
    Image4_MouseUp 1, 0, 0, 0
     Update
Else
If LF1.ListCount > 0 Then
Pid = Pid - 1
  If Pid < 0 Then Pid = LF1.ListCount - 1
MP1.Filename = LF1.List(Pid)
LF1.ListIndex = Pid
End If
End If
End If
If Button = 2 Then PopupMenu Me.b002, 0, X + 45, Y + Fm10.Top
End Sub

Private Sub Image2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If Cd = True Then
Dim E As String * 40
SendMCIString "set cd time format milliseconds", True
mciSendString "status cd position wait", E, Len(E), 0
If (Playing) Then
    Command = "play cd from " & CStr(CLng(E) - FastForwardSpeed * 500)
Else
    Command = "seek cd to " & CStr(CLng(E) - FastForwardSpeed * 500)
End If
mciSendString Command, 0, 0, 0
SendMCIString "set cd time format tmsf", True
Update
Else
If Len(MP1.Filename) > 0 Then MP1.CurrentPosition = MP1.CurrentPosition + 5
End If
End If
If Button = 2 Then PopupMenu Me.b002, 0, X + 405, Y + Fm10.Top
End Sub

Private Sub Image3_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If Cd = True Then
SendMCIString "play cd", True
Playing = True
Ig11.Picture = LoadPicture(App.Path + "\SmM_Icos\05.gif")
Exit Sub
Else
If Len(MP1.Filename) > 0 Then
MP1.Play
Exit Sub
End If
End If
If Len(MP1.Filename) = 0 And Playing = False Then
If Ly.FileExists(CDRom + "MPEGAV\AVSEQ01.DAT") = True Or Ly.FileExists(CDRom + "MPEGAV\MUSIC01.DAT") = True Then
a00106_Click
Exit Sub
End If
End If
If Len(MP1.Filename) = 0 And Playing = False Then
 Update
 If CDLoad = True Then a00105_Click
Exit Sub
End If
If Len(MP1.Filename) = 0 And Playing = False Then
If LF1.ListCount > 0 Then
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End If
End If
End If
If Button = 2 Then PopupMenu Me.b002, 0, X + 765, Y + Fm10.Top
End Sub

Private Sub Image4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If Cd = True Then
SendMCIString "pause cd", True
Playing = False
Update
Else
If Len(MP1.Filename) > 0 Then MP1.Pause
End If
End If
If Button = 2 Then PopupMenu Me.b002, 0, X + 1260, Y + Fm10.Top
End Sub

Private Sub Image5_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If Cd = True Then
SendMCIString "stop cd wait", True
Command = "seek cd to " & 1
SendMCIString Command, True
Playing = False
Update
Cd = False
Else
 If Len(MP1.Filename) > 0 Then
 SetBookMark
 MP1.Filename = "ilxz"
End If
End If
Ig11.Picture = LoadPicture(App.Path + "\SmM_Icos\08.gif")
End If
If Button = 2 Then PopupMenu Me.b002, 0, X + 1620, Y + Fm10.Top
End Sub

Private Sub Image6_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If Cd = True Then
Dim E As String * 40
SendMCIString "set cd time format milliseconds", True
mciSendString "status cd position wait", E, Len(E), 0
If (Playing) Then
    Command = "play cd from " & CStr(CLng(E) + FastForwardSpeed * 500)
Else
    Command = "seek cd to " & CStr(CLng(E) + FastForwardSpeed * 500)
End If
mciSendString Command, 0, 0, 0
SendMCIString "set cd time format tmsf", True
Update
Else
If Len(MP1.Filename) > 0 Then MP1.CurrentPosition = MP1.CurrentPosition + 5
End If
End If
If Button = 2 Then PopupMenu Me.b002, 0, X + 2475, Y + Fm10.Top
End Sub

Private Sub Image7_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Image8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If Cd = True Then
If (Track < TotalTracks) Then
    If (Playing) Then
        Command = "play cd from " & Track + 1
        SendMCIString Command, True
    Else
        Command = "seek cd to " & Track + 1
        SendMCIString Command, True
    End If
    Command = "seek cd to " & Track + 1
        SendMCIString Command, True
 Update
 Image4_MouseUp 1, 0, 0, 0
Else
If (Playing) Then
        Command = "play cd from " & 1
        SendMCIString Command, True
    Else
        Command = "seek cd to " & 1
        SendMCIString Command, True
    End If
    Command = "seek cd to " & 1
        SendMCIString Command, True
 Update
 Image4_MouseUp 1, 0, 0, 0
    Exit Sub
End If
Else
If LF1.ListCount > 0 Then
Pid = Pid + 1
  If Pid > LF1.ListCount - 1 Then Pid = 0
MP1.Filename = LF1.List(Pid)
LF1.ListIndex = Pid
End If
End If
End If
If Button = 2 Then PopupMenu Me.b002, 0, X + 2835, Y + Fm10.Top
End Sub

Private Sub Image8_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub Image9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PopupMenu Me.asdf, 0, Image9.Left + X, Fm10.Top + Y
If Button = 2 Then PopupMenu Me.b002, 0, X + 3240, Y + Fm10.Top
End Sub

Private Sub Image9_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub jhl_Click()
Image6_MouseUp 1, 0, 0, 0
End Sub

Private Sub jkjhl_Click()
Image4_MouseUp 1, 0, 0, 0
End Sub

Private Sub jlk_Click()
Image2_MouseUp 1, 0, 0, 0
End Sub

Private Sub jlkl_Click()
Image5_MouseUp 1, 0, 0, 0
End Sub

Public Sub juyjk_Click()
  clt1_ItemClick 1, 1
End Sub

Private Sub juyjuk_Click()
a00202_Click
End Sub

Private Sub kuykuk_Click()
a002020_Click
End Sub

Private Sub lf1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim pos As Long, idx As Long
pos = X / Screen.TwipsPerPixelX + Y / Screen.TwipsPerPixelY * 65536
idx = SendMessage(LF1.hWnd, LB_ITEMFROMPOINT, 0, ByVal pos)
' idx 即等于鼠标所在位置的选项
If idx < 65536 Then
LF1.ListIndex = idx
LF1.ToolTipText = "[" + Str(idx + 1) + " -" + Str(LF1.ListCount) + " ] " + LF1.List(idx)
End If
End Sub

Sub LoadSet()
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Bar_A") = True Then
Cl0.DisplayStyle = 1
Cl0.InsCharacter = "_"
End If
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Bar_B") = True Then
Cl0.DisplayStyle = 4
Cl0.InsCharacter = " "
End If
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Bar_C") = True Then
Cl0.DisplayStyle = 2
Cl0.InsCharacter = " "
End If
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Bar_D") = True Then
Cl0.DisplayStyle = 4
Cl0.InsCharacter = ""
End If
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "MousePu") = True Then
MP1.ClickToPlay = True
Else
MP1.ClickToPlay = False
End If
BSv
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ShowSys") = True Then
ST1.Action = 0
Else: ST1.Action = 2
End If
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Change", False
End Sub
Sub BSv()
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ScSave") = True Then
       If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "OnlyVideo") = True Then
             If MP1.ImageSourceWidth > 0 Then
                      DbSv
                  Else:
                      EbSv
             End If
       Else
             DbSv
       End If
Else
  EbSv
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
'Ly.Addhorizon LF1, 500
If App.PrevInstance = True Then End
Dim text As String
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Sting", True
RushBm
Ig8.Picture = LoadPicture(App.Path + "\SmM_Skin\Snowman.gif")
Ly.CenterForm Me
FastForwardSpeed = 10
CDLoad = False
SendMCIString "open cdaudio alias cd wait shareable", True
SendMCIString "set cd time format tmsf wait", True
    Open App.Path + "\SmM_Start.dat" For Input As #1
    While Not EOF(1)
    Line Input #1, text
    LF1.AddItem RTrim(text)
    Wend
    Close #1
   Open App.Path + "\SmM_Start.dat" For Output As #1
    For i = 0 To LF1.ListCount - 1
     Print #1, LeftB(LF1.List(i), 2000)
    Next i
   Close (1)
 LF1.Clear
 Open App.Path + "\SmM_List.sml" For Input As #1
    While Not EOF(1)
    Line Input #1, text
    LF1.AddItem RTrim(text)
    Wend
    Close #1
    Pid = 0
   If LF1.ListCount > 0 Then LF1.ListIndex = Pid
SF1.SkinPath = App.Path + "\SmM_Skin"
   cLT1.AddListImage 1, "S.flake.", 4
    cLT1.AddListImage 1, "打开媒体", 2
   cLT1.AddListImage 1, "当前播放", 5
    cLT1.AddListImage 1, "外观视图", 1
    cLT1.AddListImage 1, "媒体书签", 7
   cLT1.AddListImage 1, "曲目列表", 6
   cLT1.AddListImage 1, "个人媒体", 3
    cLT1.AddListImage 1, "更改选项", 9
   cLT1.AddListImage 1, "帮助支持", 8
  Cd = False
Ig11.Picture = LoadPicture(App.Path + "\SmM_Icos\08.gif")
Ig12.Picture = LoadPicture(App.Path + "\SmM_Icos\10.gif")
Ig13.Picture = LoadPicture(App.Path + "\SmM_Icos\02.gif")
          CommonDialog1.Filter = "媒体文件 (多种被支持的类型)" & _
          "|*.smm;*.sma;*.smv;*.sml;*.ilxz;*.asf;*.asx;*.wm;*.wmx;*.wmp;*.wma;*.wax;*.wmv;*.wvx;*.vob;*.cda;*.wav;*.avi;*.mpeg;*.mpg;*.mpe;*.m1v;*.mp2;*.mpv2;*.mp2v;*.mpa;*.mp3;*.m3u;*.mid;*.midi;*.rmi;*.ivf;*.aif;*.aifc;*.aiff;*.au;*.snd;*.swf|图片文件 (*.bmp;*.jpg;*.gif)|*.bmp;*jpg;*.gif|所有文件 (*.*)|*.*"
          CommonDialog1.FilterIndex = 1
File1.Pattern = "*.smm;*.sma;*.smv;*.sml;*.dat;*.ilxz;*.asf;*.asx;*.wm;*.wmx;*.wmp;*.wma;*.wax;*.wmv;*.wvx;*.vob;*.cda;*.wav;*.avi;*.mpeg;*.mpg;*.mpe;*.m1v;*.mp2;*.mpv2;*.mp2v;*.mpa;*.mp3;*.m3u;*.mid;*.midi;*.rmi;*.ivf;*.aif;*.aifc;*.aiff;*.au;*.snd;*.swf"
  CDRom = Left(Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "CDRom"), 3)
    Dim l As Long
    Dim wHotkey As Long
    wHotkey = (HOTKEYF_ALT Or HOTKEYF_CONTROL) * 256 + vbKeyS
    l = SendMessage(Me.hWnd, WM_SETHOTKEY, wHotkey, 0)
ST1.Icon = App.Path + "\SmM_Icons\003.ico"



















Me.Top = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LT")
Me.Left = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LL")
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Mt") = 1 Then a00308_Click
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Sf") = 1 Then a003012_Click
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Cp") = 1 Then a00201_Click
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Au") = 1 Then a002045_Click
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Aw") = 1 Then a00205_Click
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_At") = 1 Then a00308_Click
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LM") = 1 Then Me.WindowState = 2

If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "VideoSize") = 1 Then mjjhjm_Click
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "VideoSize") = 2 Then cerde_Click
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "VideoSize") = 3 Then cbcxbf_Click
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "VideoSize") = 4 Then xfg_Click
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "VideoSize") = 5 Then a00306_Click
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Cool") = 1 Then a003010_Click
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Skin") = 1 Then juyjk_Click
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Ftv") = 1 Then fghgt_Click

















LoadSet
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoMedia") = True Then Shell (App.Path + "\SmM_Types.exe")
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoStart") = True Then MP1.Filename = App.Path + "\SmM_Medias\Ftso.asx"
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoCnt") = True Then s_Click
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "StrCln") = True And Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoCln") = True Then a005052_Click


End Sub


Private Sub mjjhjm_Click()
MP1.DisplaySize = mpFitToSize
mjjhjm.Checked = True
cerde.Checked = False
cbcxbf.Checked = False
a00306.Checked = False
xfg.Checked = False
erg.Checked = False
End Sub

Private Sub MP1_DblClick(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
a00306_Click
End Sub

Private Sub MP1_MouseDown(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub MP1_MouseMove(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Form1.Left = Form1.Left + (X - MoveX)
Form1.Top = Form1.Top + (Y - MoveY)
End Sub

Private Sub MP1_MouseUp(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X + MP1.Left + Fm0.Left, Y + MP1.Top + Fm0.Top
End Sub

Private Sub rg_Click()
a00701_Click
End Sub

Private Sub rthrtb_Click()
hgj_Click
End Sub

Private Sub rty_Click()
HB.SetVolume 47, 47, 0
bfgb.Checked = False
 bfgg.Checked = False
 bfrg.Checked = False
 bcvhbfg.Checked = False
  vcbfg.Checked = False
 fghrt.Checked = False
 cvb.Checked = False
  rty.Checked = True
  hgj.Checked = False
  bsrety.Checked = False
 ertgre.Checked = False
 dfhgreff.Checked = False
  sedgghh.Checked = False
 vnrtt.Checked = False
  asergrjh.Checked = False
 brthsr.Checked = False
 brt.Checked = True
 rthrtb.Checked = False
End Sub

Private Sub s_Click()
  LF1.Clear
 If Ly.FileExists(myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_X", "")) = True Then
     MP1.Filename = myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_X", "")
     MP1.CurrentPosition = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_X")
     LF1.AddItem MP1.Filename, 0
   Else
     MsgBox "找不到媒体文件,文件可能已被移动或改名。", vbExclamation
   End If
End Sub

Private Sub sdf_Click()
a00101_Click
End Sub

Private Sub sdfg_Click()
clt1_ItemClick 1, 1
End Sub

Private Sub sdtg56_Click()
Shell App.Path + "\SmM_Start.exe", vbNormalFocus
End Sub

Private Sub sedgghh_Click()
bcvhbfg_Click
End Sub

Private Sub SF1_OnSkinNotify(ByVal SkinClass As String, ByVal SkinEvent As String)
    Select Case SkinClass
    Case "A"
         Image4_MouseUp 1, 0, 0, 0
    Case "B"
         Image5_MouseUp 1, 0, 0, 0
    Case "C"
         Image2_MouseUp 1, 0, 0, 0
    Case "D"
         Image8_MouseUp 1, 0, 0, 0
    Case "E"
         Image3_MouseUp 1, 0, 0, 0
    Case "F"
         Image6_MouseUp 1, 0, 0, 0
    Case "G"
         Image10_MouseUp 1, 0, 0, 0
    Case "H"
         Image7_MouseUp 1, 0, 0, 0
    Case "I"
         Image9_MouseUp 1, 0, 0, 0
    'Case "J"
         'Call Image1700_Click
    Case "min"
         Me.WindowState = 1
    Case "all"
          clt1_ItemClick 1, 1
    End Select
End Sub

Private Sub Form_Resize()
If TimeOne = False Then
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Mt") = 1 Then a00308_Click
TimeOne = True
End If
If juyjk.Checked = True Then Exit Sub
If Me.WindowState <> 1 Then
If Me.Width < 7605 Then Me.Width = 7605
If Me.Height < 5205 Then Me.Height = 5205
Fm0.Visible = False
Fm0.Width = Me.Width
Fm0.Height = Me.Height
Ig0.Width = Me.Width
Ig1.Width = Me.Width
Ig2.Width = Me.Width
Fm10.Width = Me.Width
'MP1.Width = Me.Width - 3180
MP1.Width = Me.Width - 3150
LF1.Left = Me.Width - 3060
'Ig10.Left = Me.Width - 2475
Ig11.Left = Me.Width - 2205
Ig12.Left = Me.Width - 1525
Ig13.Left = Me.Width - 845
Frame1.Top = Me.Height - 825
Frame1.Width = Me.Width - 3300
TrackTime.Left = Frame1.Width - 2535
TotalTrack.Left = Frame1.Width - 1275
cLT1.Left = Me.Width - 3060
Fm5.Left = Me.Width - 3135
LF1.Height = Me.Height - 2115
MP1.Height = Me.Height - 915
Ig7.Width = Me.Width - 3150
Ig7.Height = Me.Height - 2550
Cl0.Width = Me.Width - 2500
Ig2.Top = Me.Height - 1875
Ig1.Top = Me.Height - 1560
Fm10.Top = Me.Height - 1290
'Ig3.Top = Me.Height - 1290
cLT1.Top = Me.Height - 1530
Fm5.Top = Me.Height - 1530
Ig8.Left = (Me.Width - 3135 + Ig8.Width) / 2 - Ig8.Width + 45
Ig8.Top = (Me.Height - 2010 + Ig8.Height) / 2 - Ig8.Height + 275
Fm0.Visible = True
End If
If Me.WindowState = 1 Then
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ShowRw") = True Then Me.Visible = False
Else
LsTy = Me.WindowState
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SetBookMark
Unload Form100
Unload Form2
Unload Form3
'a003010.Checked = False
'ythg.Checked = False
Image6_MouseUp 1, 0, 0, 0
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Clean") = True Then LF1.Clear
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "OvrCln") = True And Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoCln") = True Then a005052_Click

Open App.Path + "\SmM_List.sml" For Output As #1
    For i = 0 To LF1.ListCount - 1
     Print #1, LF1.List(i)
    Next i
   Close (1)

Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Sting", False
ST1.Action = 2
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "StartUp") = True Then Shell (App.Path + "\SmM_Helper.exe")
 Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LT", Me.Top
 Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LL", Me.Left
If Me.WindowState = 2 Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LM", 1
Else
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LM", 0
End If
If a00308.Checked = True Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Mt", 1
Else
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Mt", 0
End If
If a003012.Checked = True Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Sf", 1
Else
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Sf", 0
End If
If a00201.Checked = True Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Cp", 1
Else
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Cp", 0
End If
If a002045.Checked = True Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Au", 1
Else
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Au", 0
End If
If a00205.Checked = True Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Aw", 1
Else
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Aw", 0
End If
If a00308.Checked = True Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_At", 1
Else
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_At", 0
End If
If fghgt.Checked = True Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Ftv", 1
Else
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Ftv", 0
End If

If mjjhjm.Checked = True Then Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "VideoSize", 1
If cerde.Checked = True Then Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "VideoSize", 2
If cbcxbf.Checked = True Then Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "VideoSize", 3
If xfg.Checked = True Then Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "VideoSize", 4
If a00306.Checked = True Then Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "VideoSize", 5
If a003010.Checked = True Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Cool", 1
Else
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Cool", 0
End If
If juyjk.Checked = True Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Skin", 1
Else
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Skin", 0
End If


End
End Sub
Private Function SendMCIString(Cmd As String, fShowError As Boolean) As Boolean
On Error Resume Next
Static rc As Long               'return code
Static errStr As String * 400
rc = mciSendString(Cmd, 0, 0, hWnd)
If (fShowError And rc <> 0) Then
    mciGetErrorString rc, errStr, Len(errStr)
    'MsgBox errStr
End If
SendMCIString = (rc = 0)
End Function
Private Sub Update()
On Error Resume Next
Static E As String * 30
mciSendString "status cd media present", E, Len(E), 0
If (CBool(E)) Then
    If (CDLoad = False) Then
        mciSendString "status cd number of tracks wait", E, Len(E), 0
        TotalTracks = CInt(Mid$(E, 1, 2))
        If (TotalTracks = 1) Then
            Exit Sub
        End If
        mciSendString "status cd length wait", E, Len(E), 0
        TotalTrack.Caption = TotalTracks & " / " & E
        ReDim TrackLength(1 To TotalTracks)
        Dim i As Integer
        For i = 1 To TotalTracks
            Command = "status cd length track " & i
            mciSendString Command, E, Len(E), 0
            TrackLength(i) = E
        Next
        Dim ts As Integer
        TrackSelection.Clear
        For ts = 1 To TotalTracks
        TrackSelection.AddItem ts
        Next ts
        TrackSelection.text = TrackSelection.List(0)
        CDLoad = True
        SendMCIString "seek cd to 1", True
    End If
     mciSendString "status cd position", E, Len(E), 0
    Track = CInt(Mid$(E, 1, 2))
    Minute = CInt(Mid$(E, 4, 2))
    Second = CInt(Mid$(E, 7, 2))
    TimeWindow.text = "[" & Format(Track, "00") & "] " & Format(Minute, "00") _
            & ":" & Format(Second, "00")
             TrackTime.Caption = TrackLength(Track)
    TrackSelection.text = TrackSelection.List(Track - 1)
      mciSendString "status cd mode", E, Len(E), 0
    Playing = (Mid$(E, 1, 7) = "playing")
    Else
     If (CDLoad = True) Then
        CDLoad = False
        Playing = False
        TrackTime.Caption = ""
        Me.TotalTrack.Caption = ""
        TimeWindow.text = ""
    End If
End If
End Sub

Private Sub LF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.a005, 0, LF1.Left + X, LF1.Top + Y
End Sub

Private Sub Lf1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
     Dim ThisFile As Variant
    For Each ThisFile In Data.Files
        LF1.AddItem ThisFile
    Next
    a005052_Click
End Sub

Private Sub LF1_DblClick()
Pid = LF1.ListIndex
 If Ly.FileExists(LF1.List(Pid)) = False Then
   MsgBox ("媒体文件可能已经丢失或被移动,无法播放。"), vbExclamation
 Exit Sub
 End If
 If Right(LF1.List(Pid), 4) = ".sml" Or Right(LF1.List(Pid), 4) = ".SML" Then
 a005052_Click
 End If
 If Right(LF1.List(Pid), 4) = ".swf" Or Right(LF1.List(Pid), 4) = ".SWF" Then Shell (App.Path + "\SmM_Flash.exe " + LF1.List(Pid)), vbNormalFocus
  If Right(LF1.List(Pid), 4) = ".cda" Or Right(LF1.List(Pid), 4) = ".CDA" Then
                   Playing = True
                Ig11.Picture = LoadPicture(App.Path + "\SmM_Icos\05.gif")
                      Cd = True
        If (Track <= TotalTracks) Then
            If (Playing) Then
                Command = "play cd from " & Val(Left(Right(LF1.List(Pid), 6), 2))
                SendMCIString Command, True
             Else
                Command = "seek cd to " & Val(Left(Right(LF1.List(Pid), 6), 2))
                SendMCIString Command, True
                SendMCIString "play cd", True
     End If
           MP1.Filename = "ilxz"
                Playing = True
                Cd = True
Ig11.Picture = LoadPicture(App.Path + "\SmM_Icos\05.gif")
 End If
Else
MP1.Filename = LF1.List(Pid)
  End If
End Sub

Private Sub MP1_EndOfStream(ByVal result As Long)
If MP1.Filename = App.Path + "\SmM_Medias\Ftso.asx" Then
Pid = 0
MP1.Filename = "ilxz"
Exit Sub
End If
Ig11.Picture = LoadPicture(App.Path + "\SmM_Icos\08.gif")
If a002045.Checked = False Then
   If Pid >= LF1.ListCount - 1 Then
    Pid = 0
         LF1.ListIndex = Pid
        If a00205.Checked = False Then
      MP1.Filename = "ilxz"
          Else
          LF1.ListIndex = Pid
              LF1_DblClick
           End If
    Else
LF1.ListIndex = Pid + 1
LF1_DblClick
 End If
Else
ReRnd:
i = Int((LF1.ListCount - 1) * Rnd)
If i <> Pid Then
Pid = i
Else
GoTo ReRnd
End If
LF1.ListIndex = Pid
LF1_DblClick
End If
BSv
End Sub

Private Sub MP1_OpenStateChange(ByVal OldState As Long, ByVal NewState As Long)
If MP1.Filename = App.Path + "\SmM_Medias\Ftso.asx" Then Exit Sub
If Playing = True Then
SendMCIString "stop cd wait", True
Command = "seek cd to " & 1
SendMCIString Command, True
Playing = False
Update
Cd = False
End If
If MP1.ImageSourceHeight = 0 Then
     Ig11.Picture = LoadPicture(App.Path + "\SmM_Icos\01.gif")
   Else
     If fghgt.Checked = True Then SetVw
    
         If Right(MP1.Filename, 4) = ".bmp" Or Right(MP1.Filename, 4) = ".BMP" Or Right(MP1.Filename, 4) = ".jpg" Or Right(MP1.Filename, 4) = ".JPG" Or Right(MP1.Filename, 4) = ".gif" Or Right(MP1.Filename, 4) = ".GIF" Or Right(MP1.Filename, 4) = ".smp" Or Right(MP1.Filename, 4) = ".SMP" Then
               Ig11.Picture = LoadPicture(App.Path + "\SmM_Icos\04.gif")
          Else
                If Right(MP1.Filename, 4) = ".DAT" Or Right(MP1.Filename, 4) = ".dat" Then
                    Ig11.Picture = LoadPicture(App.Path + "\SmM_Icos\12.gif")
                        Else
                        Ig11.Picture = LoadPicture(App.Path + "\SmM_Icos\06.gif")
                  End If
           End If
End If
BSv
End Sub

Private Sub ST1_MouseUp(ByVal Button As Integer)
 Me.WindowState = LsTy
 Me.Visible = True
If Button = 2 Then
 PopupMenu Me.b002, 0, Screen.Width, Screen.Height
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If MP1.Filename = App.Path + "\SmM_Medias\Ftso.asx" Then
Frame1.Visible = False
Exit Sub
End If
Update
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Change") = True Then LoadSet
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AddFile") = True Then
LF1.AddItem myReadINI(App.Path + "\SmM_Start.dat", "Start", "AddFile", ""), LF1.ListCount
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AddFile", False
a005052_Click
End If
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "PlayFile") = True Then
   LF1.Clear
   LF1.AddItem myReadINI(App.Path + "\SmM_Start.dat", "Start", "PlayFile", ""), 0
   Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "PlayFile", False
   LF1.ListIndex = 0
   LF1_DblClick
  End If
If Cd = True Then
Frame1.Visible = True
Else: Frame1.Visible = False
End If
If MP1.DisplaySize = mpFitToSize Then
mjjhjm.Checked = True
a00306.Checked = False
erg.Checked = False
End If
Info = ""
If Len(MP1.Filename) > 0 Then
Info = "[" + Str(Pid + 1) + " -" + Str(LF1.ListCount) + " ] "
If Len(MP1.GetMediaInfoString(mpClipTitle)) > 0 Then Info = Info + "标题:" + MP1.GetMediaInfoString(mpClipTitle) + "  "
If Len(MP1.GetMediaInfoString(mpClipAuthor)) > 0 Then Info = Info + "艺术家:" + MP1.GetMediaInfoString(mpClipAuthor) + "  "
If Len(MP1.GetMediaInfoString(mpClipCopyright)) > 0 Then Info = Info + "版权:" + MP1.GetMediaInfoString(mpClipCopyright) + "  "
If Len(MP1.GetMediaInfoString(mpClipDescription)) > 0 Then Info = Info + "描述:" + MP1.GetMediaInfoString(mpClipDescription) + "  "
If MP1.Bandwidth > 0 Then Info = Info + "正在播放:" + Str(Int(MP1.Bandwidth / 1000)) + " 千字节每秒  "
If MP1.ImageSourceWidth > 0 Then
    Info = Info + "视频:" + Str(MP1.ImageSourceWidth) + " x" + Str(MP1.ImageSourceHeight) + " @" + Ly.GetDisplay + "  "
Else
    Info = Info + "仅含音频  "
End If
Info = Info + "地址:" + MP1.GetMediaInfoString(mpClipFilename) + "  - Snowman Media"
End If
If Playing = True Then
 Info = "[ 音频 CD @ 驱动器:" + Left(CDRom, 1) + " " + Str(Track) + "/" + TrackTime.Caption + " - " + TotalTrack.Caption + " ]  标题:未知标题  艺术家:未知艺术家  唱片集:未知唱片集  标识:" + Ly.GetDiskNumber(CDRom) + "  - Snowman Media"
 BSv
End If
 If Len(MP1.Filename) = 0 And Playing = False Then Info = "Enjoy your multimedia by using Snowman Media"
If Cl0.Display <> Info Then
Cl0.Display = Info
If Info = "Enjoy your multimedia by using Snowman Media" Then
ST1.SysTrayText = "Snowman Media ilxz"
Else
ST1.SysTrayText = LeftB(Info, 60) + " … Snowman Media"
End If
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ShowSys") = True Then
ST1.Action = 2
ST1.Action = 0
End If
End If
End Sub

Private Sub clt1_ItemClick(ByVal nList As Integer, ByVal nItem As Integer)
Select Case nList
        Case 1
          Select Case nItem
          Case 1
          Me.Visible = False
          If juyjk.Checked = False Then
          juyjk.Checked = True
          SF1.SkinPath = Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Path")
          MP1.ShowControls = False
          MP1.ShowStatusBar = False
          Cl0.Visible = False
          Ig0.Visible = False
          Ig1.Visible = False
          Ig11.Visible = False
          Ig12.Visible = False
          Ig13.Visible = False
          Ig2.Visible = False
          Fm10.Visible = False
          Fm5.Visible = False
          cLT1.Visible = False
          Fm0.Top = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vy")
          Fm0.Left = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vx")
          Fm0.Width = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vw")
          Fm0.Height = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vh")
          Me.Width = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_w")
          Me.Height = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_h")
          LF1.Visible = False
          MP1.Top = 0
          MP1.Left = 0
          MP1.Width = Fm0.Width
          MP1.Height = Fm0.Height
          Ig7.Visible = False
          Ig8.Top = 0
          Ig8.Left = 0
          Ig8.Picture = LoadPicture(Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Bp"))
              Else
          juyjk.Checked = False
          SF1.SkinPath = App.Path + "\SmM_Skin"
          MP1.ShowControls = True
          MP1.ShowStatusBar = True
          Cl0.Visible = True
          Ig0.Visible = True
          Ig1.Visible = True
          Ig11.Visible = True
          Ig12.Visible = True
          Ig13.Visible = True
          Ig2.Visible = True
          Fm10.Visible = True
          Fm5.Visible = True
          cLT1.Visible = True
          Fm0.Top = 0
          Fm0.Left = 0
          Me.Width = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_w")
          Me.Height = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_h")
          LF1.Visible = True
          MP1.Top = 405
          MP1.Left = 45
          Ig7.Visible = True
          Ig8.Picture = LoadPicture(App.Path + "\SmM_Skin\Snowman.gif")
          Form_Resize
      End If
          Me.Visible = True
          Case 2
           PopupMenu Me.a001, 0, cLT1.Left + 350, cLT1.Top + 800
           Case 3
            PopupMenu Me.a002, 0, cLT1.Left + 350, cLT1.Top + 800
          Case 4
           PopupMenu Me.a003, 0, cLT1.Left + 350, cLT1.Top + 800
           Case 5
            PopupMenu Me.a004, 0, cLT1.Left + 350, cLT1.Top + 800
          Case 6
           PopupMenu Me.a005, 0, cLT1.Left + 350, cLT1.Top + 800
           Case 7
            PopupMenu Me.a006, 0, cLT1.Left + 350, cLT1.Top + 800
          Case 8
           PopupMenu Me.a007, 0, cLT1.Left + 350, cLT1.Top + 800
           Case 9
            PopupMenu Me.a008, 0, cLT1.Left + 350, cLT1.Top + 800
   End Select
End Select
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
        LF1.AddItem entry
    Next ind
  End If
  If BackUp <> "" Then         ' If there is a superior
      Dir1.Path = BackUp    ' directory, move to it.
  End If
  Exit Function
DirDriverHandler:
  If Err = 7 Then         ' If Out of Memory, assume listbox just got full.
    DirDiver = True       ' Create Msg$ and set return value AbandonSearch.
  '  MsgBox "You've filled the listbox. Search being abandoned..."
    Exit Function         ' Note that EXIT procedure resets ERR to 0.
  Else                    ' Otherwise display error message and quit.
    'MsgBox Error
    End
  End If
End Function

Private Sub TimeWindow_GotFocus()
cLT1.SetFocus
End Sub

Private Sub TotalTrack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub TotalTrack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Form1.Left = Form1.Left + (X - MoveX)
Form1.Top = Form1.Top + (Y - MoveY)
End Sub

Private Sub TotalTrack_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick

End Sub

Private Sub TrackTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub TrackTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button <> 1 Then Exit Sub
Form1.Left = Form1.Left + (X - MoveX)
Form1.Top = Form1.Top + (Y - MoveY)
End Sub

Private Sub TrackTime_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
For Each ThisFile In Data.Files
LF1.Clear
LF1.AddItem ThisFile
Next
Pid = 0
LF1.ListIndex = Pid
LF1_DblClick
End Sub

Private Sub vcbfg_Click()
HB.SetVolume 28, 28, 0
bfgb.Checked = False
 bfgg.Checked = False
 bfrg.Checked = False
 bcvhbfg.Checked = False
  vcbfg.Checked = True
 fghrt.Checked = False
 cvb.Checked = False
  rty.Checked = False
  hgj.Checked = False
  bsrety.Checked = False
 ertgre.Checked = False
 dfhgreff.Checked = False
  sedgghh.Checked = False
 vnrtt.Checked = True
  asergrjh.Checked = False
 brthsr.Checked = False
 brt.Checked = False
 rthrtb.Checked = False
End Sub

Private Sub vnrtt_Click()
vcbfg_Click
End Sub

Private Sub werewr_Click()
If LF1.ListCount > 0 And LF1.SelCount > 0 Then
If MP1.Filename = LF1.List(LF1.ListIndex) Then MP1.Filename = "ilxz"
Ly.DelFile LF1.List(LF1.ListIndex)
If Ly.FileExists(LF1.List(LF1.ListIndex)) = False Then
LF1.RemoveItem (LF1.ListIndex)
If Pid <= LF1.ListIndex Then Pid = Pid - 1
End If
End If
End Sub

Private Sub xfg_Click()
MP1.DisplaySize = mpDoubleSize
mjjhjm.Checked = False
cerde.Checked = False
cbcxbf.Checked = False
a00306.Checked = False
xfg.Checked = True
erg.Checked = False
End Sub

Private Sub y5665y_Click()
a00811_Click
End Sub

Private Sub ythg_Click()
a003010_Click
End Sub

Private Sub ytjuytjkuy_Click()
Image6_MouseUp 1, 0, 0, 0
End Sub

Private Sub ytu5gf_Click()
Image2_MouseUp 1, 0, 0, 0
End Sub
