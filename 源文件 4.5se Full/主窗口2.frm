VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "SmM_Snowflake.ocx"
Object = "{244E6785-6684-11D2-943F-A976CFB4FC0C}#1.0#0"; "SmM_Lstbar.ocx"
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "SmM_Tools.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{38943DFD-2C76-11D5-8FCF-A3B833033124}#1.0#0"; "SmM_SouCtrl.ocx"
Object = "{CFB094A6-8FF0-4EF7-A644-ED122CC38E57}#1.0#0"; "SmM_Tray.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Snowman Media 4se"
   ClientHeight    =   4395
   ClientLeft      =   2025
   ClientTop       =   1815
   ClientWidth     =   7485
   Icon            =   "主窗口2.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4395
   ScaleWidth      =   7485
   Begin HYZ声音控制控件.HYZVolBan HB 
      Height          =   465
      Left            =   4635
      TabIndex        =   20
      Top             =   6480
      Visible         =   0   'False
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   820
   End
   Begin TASKICONLib.TaskIcon St1 
      Left            =   7155
      Top             =   7155
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "主窗口2.frx":2CFA
      ToolTipText     =   "Snowman Media ilxz 4"
      TitleText       =   ""
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6525
      Top             =   6525
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "浏览媒体文件"
   End
   Begin ACTIVESKINLibCtl.SkinForm Sf1 
      Height          =   480
      Left            =   5895
      OleObjectBlob   =   "主窗口2.frx":3094
      TabIndex        =   10
      Top             =   6480
      Width           =   480
   End
   Begin VB.FileListBox File1 
      Height          =   450
      Left            =   5940
      TabIndex        =   9
      Top             =   7200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.DirListBox Dir1 
      Height          =   510
      Left            =   5265
      TabIndex        =   8
      Top             =   7155
      Visible         =   0   'False
      Width           =   555
   End
   Begin API控制大全.LyfTools Ly 
      Left            =   5310
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   7065
      Top             =   6525
   End
   Begin VB.ListBox Lf1 
      Appearance      =   0  'Flat
      Height          =   390
      ItemData        =   "主窗口2.frx":30E7
      Left            =   6570
      List            =   "主窗口2.frx":30E9
      TabIndex        =   7
      Top             =   7200
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Frame Fm0 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   5460
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   0
      Width           =   7965
      Begin VB.Frame Fm5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C08062&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   4500
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   3690
         Width           =   105
         Begin VB.Image Ig 
            Appearance      =   0  'Flat
            Height          =   1125
            Index           =   9
            Left            =   0
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口2.frx":30EB
            Top             =   0
            Width           =   150
         End
      End
      Begin CTLISTBARLibCtl.ctListBar cLT1 
         Height          =   1275
         Left            =   4635
         TabIndex        =   3
         ToolTipText     =   "功能菜单"
         Top             =   3645
         Width           =   2760
         _Version        =   65536
         _ExtentX        =   4868
         _ExtentY        =   2249
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
         BackImage       =   "主窗口2.frx":3172
         BorderColor     =   12615778
         ButtonBackColor =   16777215
         ButtonForeColor =   12615778
         ListBackColor   =   12615778
         ListForeColor   =   16711680
         BarBackColor    =   12648384
         BarForeColor    =   12615778
         BorderType      =   1
         ListBarStyle    =   0
         BarHeight       =   0
         WordWrap        =   -1  'True
         Caption         =   ""
         PicArray0       =   "主窗口2.frx":3D96
         PicArray1       =   "主窗口2.frx":4670
         PicArray2       =   "主窗口2.frx":4F4A
         PicArray3       =   "主窗口2.frx":5824
         PicArray4       =   "主窗口2.frx":6876
         PicArray5       =   "主窗口2.frx":7150
         PicArray6       =   "主窗口2.frx":81A2
         PicArray7       =   "主窗口2.frx":91F4
         PicArray8       =   "主窗口2.frx":9ACE
         PicArray9       =   "主窗口2.frx":A3A8
         PicArray10      =   "主窗口2.frx":A3C4
         PicArray11      =   "主窗口2.frx":A3E0
         PicArray12      =   "主窗口2.frx":A3FC
         PicArray13      =   "主窗口2.frx":A418
         PicArray14      =   "主窗口2.frx":A434
         PicArray15      =   "主窗口2.frx":A450
         PicArray16      =   "主窗口2.frx":A46C
         PicArray17      =   "主窗口2.frx":A488
         PicArray18      =   "主窗口2.frx":A4A4
         PicArray19      =   "主窗口2.frx":A4C0
         PicArray20      =   "主窗口2.frx":A4DC
         PicArray21      =   "主窗口2.frx":A4F8
         PicArray22      =   "主窗口2.frx":A514
         PicArray23      =   "主窗口2.frx":A530
         PicArray24      =   "主窗口2.frx":A54C
         PicArray25      =   "主窗口2.frx":A568
         PicArray26      =   "主窗口2.frx":A584
         PicArray27      =   "主窗口2.frx":A5A0
         PicArray28      =   "主窗口2.frx":A5BC
         PicArray29      =   "主窗口2.frx":A5D8
         PicArray30      =   "主窗口2.frx":A5F4
         PicArray31      =   "主窗口2.frx":A610
         PicArray32      =   "主窗口2.frx":A62C
         PicArray33      =   "主窗口2.frx":A648
         PicArray34      =   "主窗口2.frx":A664
         PicArray35      =   "主窗口2.frx":A680
         PicArray36      =   "主窗口2.frx":A69C
         PicArray37      =   "主窗口2.frx":A6B8
         PicArray38      =   "主窗口2.frx":A6D4
         PicArray39      =   "主窗口2.frx":A6F0
         PicArray40      =   "主窗口2.frx":A70C
         PicArray41      =   "主窗口2.frx":A728
         PicArray42      =   "主窗口2.frx":A744
         PicArray43      =   "主窗口2.frx":A760
         PicArray44      =   "主窗口2.frx":A77C
         PicArray45      =   "主窗口2.frx":A798
         PicArray46      =   "主窗口2.frx":A7B4
         PicArray47      =   "主窗口2.frx":A7D0
         PicArray48      =   "主窗口2.frx":A7EC
         PicArray49      =   "主窗口2.frx":A808
         PicArray50      =   "主窗口2.frx":A824
         PicArray51      =   "主窗口2.frx":A840
         PicArray52      =   "主窗口2.frx":A85C
         PicArray53      =   "主窗口2.frx":A878
         PicArray54      =   "主窗口2.frx":A894
         PicArray55      =   "主窗口2.frx":A8B0
         PicArray56      =   "主窗口2.frx":A8CC
         PicArray57      =   "主窗口2.frx":A8E8
         PicArray58      =   "主窗口2.frx":A904
         PicArray59      =   "主窗口2.frx":A920
         PicArray60      =   "主窗口2.frx":A93C
         PicArray61      =   "主窗口2.frx":A958
         PicArray62      =   "主窗口2.frx":A974
         PicArray63      =   "主窗口2.frx":A990
         PicArray64      =   "主窗口2.frx":A9AC
         PicArray65      =   "主窗口2.frx":A9C8
         PicArray66      =   "主窗口2.frx":A9E4
         PicArray67      =   "主窗口2.frx":AA00
         PicArray68      =   "主窗口2.frx":AA1C
         PicArray69      =   "主窗口2.frx":AA38
         PicArray70      =   "主窗口2.frx":AA54
         PicArray71      =   "主窗口2.frx":AA70
         PicArray72      =   "主窗口2.frx":AA8C
         PicArray73      =   "主窗口2.frx":AAA8
         PicArray74      =   "主窗口2.frx":AAC4
         PicArray75      =   "主窗口2.frx":AAE0
         PicArray76      =   "主窗口2.frx":AAFC
         PicArray77      =   "主窗口2.frx":AB18
         PicArray78      =   "主窗口2.frx":AB34
         PicArray79      =   "主窗口2.frx":AB50
         PicArray80      =   "主窗口2.frx":AB6C
         PicArray81      =   "主窗口2.frx":AB88
         PicArray82      =   "主窗口2.frx":ABA4
         PicArray83      =   "主窗口2.frx":ABC0
         PicArray84      =   "主窗口2.frx":ABDC
         PicArray85      =   "主窗口2.frx":ABF8
         PicArray86      =   "主窗口2.frx":AC14
         PicArray87      =   "主窗口2.frx":AC30
         PicArray88      =   "主窗口2.frx":AC4C
         PicArray89      =   "主窗口2.frx":AC68
         PicArray90      =   "主窗口2.frx":AC84
         PicArray91      =   "主窗口2.frx":ACA0
         PicArray92      =   "主窗口2.frx":ACBC
         PicArray93      =   "主窗口2.frx":ACD8
         PicArray94      =   "主窗口2.frx":ACF4
         PicArray95      =   "主窗口2.frx":AD10
         PicArray96      =   "主窗口2.frx":AD2C
         PicArray97      =   "主窗口2.frx":AD48
         PicArray98      =   "主窗口2.frx":AD64
         PicArray99      =   "主窗口2.frx":AD80
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   45
         TabIndex        =   17
         Top             =   3690
         Width           =   4515
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            Height          =   180
            Left            =   90
            TabIndex        =   19
            Top             =   25
            Width           =   180
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   80
            ScaleHeight     =   120
            ScaleWidth      =   4335
            TabIndex        =   18
            Top             =   60
            Width           =   4335
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   90
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         Top             =   4365
         Width           =   4245
         Begin VB.Label Ct 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "准备就绪"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   360
            OLEDropMode     =   1  'Manual
            TabIndex        =   16
            Top             =   45
            Width           =   1545
         End
         Begin VB.Label Times 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   2430
            OLEDropMode     =   1  'Manual
            TabIndex        =   15
            Top             =   45
            Width           =   1860
         End
      End
      Begin VB.PictureBox Pinfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   45
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   255
         ScaleWidth      =   5115
         TabIndex        =   12
         Top             =   30
         Width           =   5145
         Begin VB.Frame Tinfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   500
            OLEDropMode     =   1  'Manual
            TabIndex        =   13
            Top             =   45
            Width           =   1680
            Begin VB.Label Txinfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   180
               Left            =   0
               OLEDropMode     =   1  'Manual
               TabIndex        =   14
               Top             =   0
               Width           =   90
            End
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   4770
         ScaleHeight     =   30
         ScaleWidth      =   2610
         TabIndex        =   11
         Top             =   1170
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.ListBox LF2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00C0FFFF&
         Height          =   2910
         IntegralHeight  =   0   'False
         ItemData        =   "主窗口2.frx":AD9C
         Left            =   4725
         List            =   "主窗口2.frx":AD9E
         OLEDropMode     =   1  'Manual
         TabIndex        =   5
         Top             =   495
         Width           =   2670
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
         TabIndex        =   4
         Top             =   3915
         Width           =   6990
         Begin VB.Image Ig 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   11
            Left            =   45
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口2.frx":ADA0
            Stretch         =   -1  'True
            ToolTipText     =   "上一首"
            Top             =   45
            Width           =   330
         End
         Begin VB.Image Ig 
            Appearance      =   0  'Flat
            Height          =   420
            Index           =   14
            Left            =   765
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口2.frx":B0E1
            Stretch         =   -1  'True
            ToolTipText     =   "播放"
            Top             =   0
            Width           =   465
         End
         Begin VB.Image Ig 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   15
            Left            =   1260
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口2.frx":B422
            Stretch         =   -1  'True
            ToolTipText     =   "暂停"
            Top             =   45
            Width           =   375
         End
         Begin VB.Image Ig 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   16
            Left            =   1620
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口2.frx":B763
            Stretch         =   -1  'True
            ToolTipText     =   "停止"
            Top             =   45
            Width           =   420
         End
         Begin VB.Image Ig 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   17
            Left            =   2475
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口2.frx":BAA4
            Stretch         =   -1  'True
            ToolTipText     =   "前进"
            Top             =   45
            Width           =   330
         End
         Begin VB.Image Ig 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   12
            Left            =   2835
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口2.frx":BDE5
            Stretch         =   -1  'True
            ToolTipText     =   "下一首"
            Top             =   45
            Width           =   375
         End
         Begin VB.Image Ig 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   18
            Left            =   3240
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口2.frx":C126
            Stretch         =   -1  'True
            ToolTipText     =   "音量"
            Top             =   45
            Width           =   375
         End
         Begin VB.Image Ig 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   13
            Left            =   405
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口2.frx":C467
            Stretch         =   -1  'True
            ToolTipText     =   "后退"
            Top             =   45
            Width           =   330
         End
         Begin VB.Image Ig 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   10
            Left            =   2070
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口2.frx":C7A8
            Stretch         =   -1  'True
            ToolTipText     =   "弹出光盘"
            Top             =   45
            Width           =   375
         End
         Begin VB.Image Ig 
            Appearance      =   0  'Flat
            Height          =   435
            Index           =   6
            Left            =   45
            OLEDropMode     =   1  'Manual
            Picture         =   "主窗口2.frx":CAE9
            Top             =   0
            Width           =   3750
         End
      End
      Begin VB.Image Ig 
         Appearance      =   0  'Flat
         Height          =   15
         Index           =   8
         Left            =   540
         OLEDropMode     =   1  'Manual
         Top             =   765
         Width           =   15
      End
      Begin VB.Image Ig 
         Height          =   510
         Index           =   19
         Left            =   3420
         MouseIcon       =   "主窗口2.frx":DC38
         MousePointer    =   99  'Custom
         Top             =   2925
         Width           =   915
      End
      Begin VB.Image Ig 
         Appearance      =   0  'Flat
         Height          =   2475
         Index           =   7
         Left            =   45
         OLEDropMode     =   1  'Manual
         Picture         =   "主窗口2.frx":DD8A
         Stretch         =   -1  'True
         Top             =   405
         Width           =   3390
      End
      Begin VB.Image Ig 
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   5
         Left            =   0
         OLEDropMode     =   1  'Manual
         Picture         =   "主窗口2.frx":E0CE
         Stretch         =   -1  'True
         Top             =   3330
         Width           =   7425
      End
      Begin MediaPlayerCtl.MediaPlayer Mper 
         Height          =   4290
         Left            =   45
         TabIndex        =   6
         Top             =   405
         Width           =   4470
         AudioStream     =   -1
         AutoSize        =   0   'False
         AutoStart       =   -1  'True
         AnimationAtStart=   0   'False
         AllowScan       =   -1  'True
         AllowChangeDisplaySize=   -1  'True
         AutoRewind      =   0   'False
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
         Volume          =   0
         WindowlessVideo =   0   'False
      End
      Begin VB.Image Ig 
         Appearance      =   0  'Flat
         Height          =   1290
         Index           =   1
         Left            =   0
         OLEDropMode     =   1  'Manual
         Picture         =   "主窗口2.frx":E518
         Stretch         =   -1  'True
         Top             =   3645
         Width           =   7590
      End
      Begin VB.Image Ig 
         Appearance      =   0  'Flat
         Height          =   15
         Index           =   2
         Left            =   5340
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "信息"
         Top             =   15
         Width           =   15
      End
      Begin VB.Image Ig 
         Appearance      =   0  'Flat
         Height          =   15
         Index           =   3
         Left            =   6015
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "原序"
         Top             =   15
         Width           =   15
      End
      Begin VB.Image Ig 
         Appearance      =   0  'Flat
         Height          =   15
         Index           =   4
         Left            =   6705
         OLEDropMode     =   1  'Manual
         ToolTipText     =   "循环"
         Top             =   15
         Width           =   15
      End
      Begin VB.Image Ig 
         Appearance      =   0  'Flat
         Height          =   510
         Index           =   0
         Left            =   0
         OLEDropMode     =   1  'Manual
         Picture         =   "主窗口2.frx":E5AB
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7665
      End
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
         Caption         =   "浏览(&B)..."
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
         Caption         =   "媒体光盘(&M)"
      End
      Begin VB.Menu a00109 
         Caption         =   "-"
      End
      Begin VB.Menu a001010 
         Caption         =   "HDTV (&H)"
      End
      Begin VB.Menu a001011 
         Caption         =   "在线调谐(&R)"
      End
      Begin VB.Menu a001012 
         Caption         =   "-"
      End
      Begin VB.Menu dfgfg 
         Caption         =   "网页(&P)"
      End
      Begin VB.Menu a001013 
         Caption         =   "媒体指南(&W)"
         Enabled         =   0   'False
      End
      Begin VB.Menu frgfrefaew 
         Caption         =   "-"
      End
      Begin VB.Menu dfg 
         Caption         =   "弹出光盘(&E)"
         Shortcut        =   ^E
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
         Caption         =   "上一首(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu fgrgfrgrgrtg 
         Caption         =   "下一首(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu greg 
         Caption         =   "后退(&A)"
      End
      Begin VB.Menu dfgr 
         Caption         =   "前进(&E)"
      End
      Begin VB.Menu gdsrg 
         Caption         =   "-"
      End
      Begin VB.Menu ccdf 
         Caption         =   "获取信息(&G)"
         Begin VB.Menu sdfewe 
            Caption         =   "依据标题(&T)"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu sdfewe 
            Caption         =   "依据艺术家(&A)"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu sdfewe 
            Caption         =   "依据唱片集(&D)"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu hjew 
            Caption         =   "-"
         End
         Begin VB.Menu sdfewer 
            Caption         =   "自定义(&M)"
         End
      End
      Begin VB.Menu fdhgrtjbvg 
         Caption         =   "同赏一曲(&T)"
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu sdfsdfefe 
         Caption         =   "-"
      End
      Begin VB.Menu a002020 
         Caption         =   "统计信息(&I)..."
      End
      Begin VB.Menu a00202 
         Caption         =   "属性(&R)..."
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
         Caption         =   "最大化(&M)"
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
      Begin VB.Menu gmh 
         Caption         =   "透明度(C)"
         Begin VB.Menu yj6 
            Caption         =   "不透明(&O)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu yj6 
            Caption         =   "10%"
            Index           =   1
         End
         Begin VB.Menu yj6 
            Caption         =   "20%"
            Index           =   2
         End
         Begin VB.Menu yj6 
            Caption         =   "30%"
            Index           =   3
         End
         Begin VB.Menu yj6 
            Caption         =   "40%"
            Index           =   4
         End
         Begin VB.Menu yj6 
            Caption         =   "50%"
            Index           =   5
         End
         Begin VB.Menu yj6 
            Caption         =   "60%"
            Index           =   6
         End
         Begin VB.Menu yj6 
            Caption         =   "70%"
            Index           =   7
         End
         Begin VB.Menu yj6 
            Caption         =   "80%"
            Index           =   8
         End
         Begin VB.Menu yj6 
            Caption         =   "90%"
            Index           =   9
         End
      End
      Begin VB.Menu fsdcfd555 
         Caption         =   "-"
      End
      Begin VB.Menu a003012 
         Caption         =   "选择 Snowflake(&H)..."
         Shortcut        =   ^C
      End
      Begin VB.Menu fdhj 
         Caption         =   "-"
      End
      Begin VB.Menu hgnmbmjh 
         Caption         =   "视频缩放(&V)"
         Begin VB.Menu fghgt 
            Caption         =   "窗口适应视频(&W)"
         End
         Begin VB.Menu mjjhjm 
            Caption         =   "视频适应窗口(&V)"
         End
         Begin VB.Menu xcdcdcd 
            Caption         =   "-"
         End
         Begin VB.Menu cerde 
            Caption         =   "50%"
         End
         Begin VB.Menu cbcxbf 
            Caption         =   "100%"
         End
         Begin VB.Menu xfg 
            Caption         =   "200%"
         End
         Begin VB.Menu dfxswefrff 
            Caption         =   "-"
         End
         Begin VB.Menu a00306 
            Caption         =   "全屏幕(&F)"
            Shortcut        =   ^X
         End
      End
      Begin VB.Menu gdfg 
         Caption         =   "可视效果(&I)"
         Enabled         =   0   'False
         Begin VB.Menu dsfsdf 
            Caption         =   ""
         End
      End
      Begin VB.Menu a00201 
         Caption         =   "字幕(&W)"
      End
   End
   Begin VB.Menu a004 
      Caption         =   "e"
      Begin VB.Menu s 
         Caption         =   "记忆播放"
         Shortcut        =   {F5}
      End
      Begin VB.Menu a00402 
         Caption         =   "-"
      End
      Begin VB.Menu a00403 
         Caption         =   ""
         Index           =   0
         Shortcut        =   +{F1}
      End
      Begin VB.Menu a00403 
         Caption         =   ""
         Index           =   1
         Shortcut        =   +{F2}
      End
      Begin VB.Menu a00403 
         Caption         =   ""
         Index           =   2
         Shortcut        =   +{F3}
      End
      Begin VB.Menu a00403 
         Caption         =   ""
         Index           =   3
         Shortcut        =   +{F4}
      End
      Begin VB.Menu a00403 
         Caption         =   ""
         Index           =   4
         Shortcut        =   +{F5}
      End
      Begin VB.Menu a00407 
         Caption         =   "-"
      End
      Begin VB.Menu a00408 
         Caption         =   "标记书签 [A]"
         Index           =   0
         Shortcut        =   +^{F1}
      End
      Begin VB.Menu a00408 
         Caption         =   "标记书签 [B]"
         Index           =   1
         Shortcut        =   +^{F2}
      End
      Begin VB.Menu a00408 
         Caption         =   "标记书签 [C]"
         Index           =   2
         Shortcut        =   +^{F3}
      End
      Begin VB.Menu a00408 
         Caption         =   "标记书签 [D]"
         Index           =   3
         Shortcut        =   +^{F4}
      End
      Begin VB.Menu a00408 
         Caption         =   "标记书签 [E]"
         Index           =   4
         Shortcut        =   +^{F5}
      End
   End
   Begin VB.Menu a005 
      Caption         =   "f"
      Begin VB.Menu a00001 
         Caption         =   "播放(&P)"
      End
      Begin VB.Menu dsfasdsafe 
         Caption         =   "属性(&R)..."
      End
      Begin VB.Menu rgfdgwqr 
         Caption         =   "-"
      End
      Begin VB.Menu dsgegd 
         Caption         =   "重命名(&N)"
      End
      Begin VB.Menu gfdsgawe 
         Caption         =   "更改连接(&I)"
      End
      Begin VB.Menu a005042 
         Caption         =   "-"
      End
      Begin VB.Menu a00002 
         Caption         =   "删除所选(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu werewr 
         Caption         =   "删除文件(&F)"
      End
      Begin VB.Menu a00003 
         Caption         =   "-"
      End
      Begin VB.Menu a00501 
         Caption         =   "添加文件(&A)..."
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu a005021 
         Caption         =   "添加地址(&E)..."
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu a005031 
         Caption         =   "添加文件夹(&O)..."
         Shortcut        =   +^{F12}
      End
      Begin VB.Menu dfasdfasdfw 
         Caption         =   "-"
      End
      Begin VB.Menu a005074 
         Caption         =   "清空(&L)"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu a005063 
         Caption         =   "导出(&X)..."
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu a006 
      Caption         =   "g"
      Begin VB.Menu a00601 
         Caption         =   "媒体库(&M)"
         Enabled         =   0   'False
         Shortcut        =   ^I
      End
      Begin VB.Menu sdafgerg 
         Caption         =   "-"
      End
      Begin VB.Menu a005010 
         Caption         =   "连接随身听(&O)"
         Enabled         =   0   'False
      End
      Begin VB.Menu a00502 
         Caption         =   "-"
      End
      Begin VB.Menu a00503 
         Caption         =   "从 CD 复制(&C)"
      End
      Begin VB.Menu sdfhgjy 
         Caption         =   "从 VCD、DVD 复制(&V)"
         Enabled         =   0   'False
      End
      Begin VB.Menu jtrhsgcdgr 
         Caption         =   "-"
      End
      Begin VB.Menu a00505 
         Caption         =   "视频捕获(&I)"
      End
      Begin VB.Menu greyhy 
         Caption         =   "-"
      End
      Begin VB.Menu a00509 
         Caption         =   "刻录 CD (&B)"
         Enabled         =   0   'False
      End
      Begin VB.Menu jjjyyy 
         Caption         =   "网络广播(&N)"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu a007 
      Caption         =   "h"
      Begin VB.Menu a00701 
         Caption         =   "选项(&O)"
         Shortcut        =   ^Q
      End
      Begin VB.Menu a00703 
         Caption         =   "播放选项(&P)..."
      End
      Begin VB.Menu gh 
         Caption         =   "-"
      End
      Begin VB.Menu fhghg 
         Caption         =   "均衡(&E)"
         Enabled         =   0   'False
      End
      Begin VB.Menu d 
         Caption         =   "混合器(&M)"
      End
      Begin VB.Menu a00705 
         Caption         =   "-"
      End
      Begin VB.Menu a00707 
         Caption         =   "许可证管理(&L)"
         Enabled         =   0   'False
      End
      Begin VB.Menu a00702 
         Caption         =   "-"
      End
      Begin VB.Menu fcv 
         Caption         =   "自动关机(&S)"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu a008 
      Caption         =   "i"
      Begin VB.Menu a00801 
         Caption         =   "帮助(&H)"
         Shortcut        =   ^H
      End
      Begin VB.Menu fghdrh 
         Caption         =   "系统信息(&S)"
      End
      Begin VB.Menu a00802 
         Caption         =   "-"
      End
      Begin VB.Menu a00803 
         Caption         =   "在线更新(&U)"
      End
      Begin VB.Menu sdfewf 
         Caption         =   "-"
      End
      Begin VB.Menu a00805 
         Caption         =   "流动网络(&I)"
         Shortcut        =   ^G
      End
      Begin VB.Menu a00806 
         Caption         =   "交流反馈(&O)"
      End
      Begin VB.Menu a00807 
         Caption         =   "-"
      End
      Begin VB.Menu a00808 
         Caption         =   "自述(&C)"
      End
      Begin VB.Menu a00809 
         Caption         =   "协议(&L)"
      End
      Begin VB.Menu a00810 
         Caption         =   "-"
      End
      Begin VB.Menu a00811 
         Caption         =   "关于 Snomwan(&A)..."
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu b002 
      Caption         =   "j"
      Begin VB.Menu sdf 
         Caption         =   "打开(&O)..."
      End
      Begin VB.Menu dfd 
         Caption         =   "加入列表(&T)..."
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
         Caption         =   "上一首(&B)"
      End
      Begin VB.Menu il 
         Caption         =   "下一首(&F)"
      End
      Begin VB.Menu sdf665y 
         Caption         =   "-"
      End
      Begin VB.Menu jhdtg 
         Caption         =   "音量(&V)"
         Begin VB.Menu asdfa 
            Caption         =   "增大(&U)"
            Index           =   0
         End
         Begin VB.Menu asdfa 
            Caption         =   "减小(&D)"
            Index           =   1
         End
         Begin VB.Menu dvsd 
            Caption         =   "-"
         End
         Begin VB.Menu rthrtb 
            Caption         =   "最大(&M)"
            Index           =   0
         End
         Begin VB.Menu rthrtb 
            Caption         =   "90%"
            Index           =   1
         End
         Begin VB.Menu rthrtb 
            Caption         =   "80%"
            Index           =   2
         End
         Begin VB.Menu rthrtb 
            Caption         =   "70%"
            Index           =   3
         End
         Begin VB.Menu rthrtb 
            Caption         =   "60%"
            Index           =   4
         End
         Begin VB.Menu rthrtb 
            Caption         =   "50%"
            Checked         =   -1  'True
            Index           =   5
         End
         Begin VB.Menu rthrtb 
            Caption         =   "40%"
            Index           =   6
         End
         Begin VB.Menu rthrtb 
            Caption         =   "30%"
            Index           =   7
         End
         Begin VB.Menu rthrtb 
            Caption         =   "20%"
            Index           =   8
         End
         Begin VB.Menu rthrtb 
            Caption         =   "10%"
            Index           =   9
         End
         Begin VB.Menu rthrtb 
            Caption         =   "静音(&N)"
            Index           =   10
         End
      End
      Begin VB.Menu fgb 
         Caption         =   "声道(&L)"
         Begin VB.Menu dgv 
            Caption         =   "立体声(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu vsdvrvxz 
            Caption         =   "-"
         End
         Begin VB.Menu zxvreg 
            Caption         =   "左声道(&L)"
         End
         Begin VB.Menu zxvrs 
            Caption         =   "右声道(&R)"
         End
      End
      Begin VB.Menu jj78 
         Caption         =   "-"
      End
      Begin VB.Menu fnbrth 
         Caption         =   "DVD 功能(&D)"
         Enabled         =   0   'False
      End
      Begin VB.Menu cvbtr 
         Caption         =   "-"
      End
      Begin VB.Menu erth 
         Caption         =   "始终前置(&W)"
      End
      Begin VB.Menu xtgf4 
         Caption         =   "透明度(&C)"
         Begin VB.Menu asd2 
            Caption         =   "不透明(&O)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu asd2 
            Caption         =   "10%"
            Index           =   1
         End
         Begin VB.Menu asd2 
            Caption         =   "20%"
            Index           =   2
         End
         Begin VB.Menu asd2 
            Caption         =   "30%"
            Index           =   3
         End
         Begin VB.Menu asd2 
            Caption         =   "40%"
            Index           =   4
         End
         Begin VB.Menu asd2 
            Caption         =   "50%"
            Index           =   5
         End
         Begin VB.Menu asd2 
            Caption         =   "60%"
            Index           =   6
         End
         Begin VB.Menu asd2 
            Caption         =   "70%"
            Index           =   7
         End
         Begin VB.Menu asd2 
            Caption         =   "80%"
            Index           =   8
         End
         Begin VB.Menu asd2 
            Caption         =   "90%"
            Index           =   9
         End
      End
      Begin VB.Menu juyjk 
         Caption         =   "Snowflake 模式(&N)"
      End
      Begin VB.Menu erg 
         Caption         =   "视频全屏(&E)"
      End
      Begin VB.Menu juyu 
         Caption         =   "-"
      End
      Begin VB.Menu vbx 
         Caption         =   "可视效果(&E)"
         Enabled         =   0   'False
         Begin VB.Menu rtghfg 
            Caption         =   ""
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu sdfxc 
         Caption         =   "-"
      End
      Begin VB.Menu csdaewwq 
         Caption         =   "获得信息(&G)"
         Begin VB.Menu dfdfvv 
            Caption         =   "依据标题(&T)"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu dfdfvv 
            Caption         =   "依据艺术家(&A)"
            Enabled         =   0   'False
            Index           =   1
         End
         Begin VB.Menu dfdfvv 
            Caption         =   "依据唱片集(&D)"
            Enabled         =   0   'False
            Index           =   2
         End
         Begin VB.Menu cdg 
            Caption         =   "-"
         End
         Begin VB.Menu veg 
            Caption         =   "自定义(&M)"
         End
      End
      Begin VB.Menu kuykuk 
         Caption         =   "统计信息(&I)..."
      End
      Begin VB.Menu juyjuk 
         Caption         =   "属性(&R)..."
      End
      Begin VB.Menu juykik 
         Caption         =   "-"
      End
      Begin VB.Menu kuyok 
         Caption         =   "选项(&M)"
         Begin VB.Menu rg 
            Caption         =   "选项(&O)"
         End
         Begin VB.Menu ghjykjyt 
            Caption         =   "播放选项(&P)..."
         End
      End
      Begin VB.Menu iuot7oi7to 
         Caption         =   "帮助(&H)"
      End
      Begin VB.Menu y5665y 
         Caption         =   "关于(&A)..."
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
      Begin VB.Menu asdfawe 
         Caption         =   "增大(U)"
         Index           =   0
         Shortcut        =   {F11}
      End
      Begin VB.Menu asdfawe 
         Caption         =   "减小(D)"
         Index           =   1
         Shortcut        =   {F12}
      End
      Begin VB.Menu sdcvd 
         Caption         =   "-"
      End
      Begin VB.Menu bfgb 
         Caption         =   "最大 (&M)"
         Index           =   0
         Shortcut        =   ^{F11}
      End
      Begin VB.Menu bfgb 
         Caption         =   "90%"
         Index           =   1
         Shortcut        =   ^{F9}
      End
      Begin VB.Menu bfgb 
         Caption         =   "80%"
         Index           =   2
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu bfgb 
         Caption         =   "70%"
         Index           =   3
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu bfgb 
         Caption         =   "60%"
         Index           =   4
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu bfgb 
         Caption         =   "50%"
         Checked         =   -1  'True
         Index           =   5
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu bfgb 
         Caption         =   "40%"
         Index           =   6
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu bfgb 
         Caption         =   "30%"
         Index           =   7
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu bfgb 
         Caption         =   "20%"
         Index           =   8
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu bfgb 
         Caption         =   "10%"
         Index           =   9
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu bfgb 
         Caption         =   "静音(&N)"
         Index           =   10
         Shortcut        =   ^N
      End
      Begin VB.Menu fgjyt 
         Caption         =   "-"
      End
      Begin VB.Menu fbgrt 
         Caption         =   "声道(&S)"
         Begin VB.Menu xbcty 
            Caption         =   "立体声(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu bxctyu 
            Caption         =   "-"
         End
         Begin VB.Menu xbrth 
            Caption         =   "左声道(&L)"
         End
         Begin VB.Menu bf6t5 
            Caption         =   "右声道(&R)"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SinfoBo As Boolean
Dim CDid As Long
Dim ReAdd As Long
Dim TX As Long

Dim Lor As Boolean
Dim Oline  As Long
Dim nItemX As Long
Dim nItemY As Long
Dim MoveX As Long
Dim MoveY As Long
Dim Stime As String
Const WM_SETHOTKEY = &H32
Const HOTKEYF_CONTROL = &H2
Const HOTKEYF_ALT = &H4
Dim Info As String
Dim CDRom As String
Dim i As Integer
Public Pid As Long
Dim cd As Boolean
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_ITEMFROMPOINT = &H1A9
Dim Playing As Boolean                ' true if CD is currently playing
Dim Track As Long                     ' current track
Dim Minute As Long                   ' current minute on track
Dim Second As Long                  ' current second on track
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
        "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As _
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
Dim SearchFlag As Long    ' Used as flag for cancelling, etc.
Const SPI_SETSCREENSAVEACTIVE = 17
Const SPI_SETSCREENSAVETIMEOUT = 15
Const SPIF_SENDWININICHANGE = &H2
Private Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
     ByVal lpvParam As Long, ByVal fuWinIni As Long) As Long
Dim LsTy As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long


Private Sub EbSv()
  Call SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, 1, 0, SPIF_SENDWININICHANGE)
End Sub

Private Sub DbSv()
  Call SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, 0, 0, SPIF_SENDWININICHANGE)
End Sub

Private Function GetFolderValue(wIdx As Long) As Long
    If wIdx < 2 Then
        GetFolderValue = 0
    ElseIf wIdx < 12 Then
        GetFolderValue = wIdx
    Else
        GetFolderValue = wIdx + 4
    End If
End Function


Private Sub a00001_Click()
LF2_DblClick
End Sub


Private Sub a00002_Click()
On Error Resume Next
If LF2.ListCount > 0 And LF2.SelCount > 0 Then
If Pid <= LF2.ListIndex Then Pid = Pid - 1
Lf1.RemoveItem (LF2.ListIndex)
LF2.RemoveItem (LF2.ListIndex)
  LF2.ListIndex = Pid
End If
End Sub

Private Sub a001010_Click()
Shell App.Path + "\SmM_HDTV.exe", vbNormalFocus
End Sub

Private Sub a001011_Click()
File1.Path = App.Path + "\SmM_NetMedias"
Lf1.Clear
LF2.Clear
For i = 0 To File1.ListCount - 1
Lf1.AddItem App.Path + "\SmM_NetMedias\" + File1.List(i)
LF2.AddItem Left(File1.List(i), Len(File1.List(i)) - 4)
Next
End Sub

Private Sub a00102_Click()
Dim SelectFileName As String
 SelectFileName = InputBox("请输入万维网地址 (URL) 或指定你要打开的本地媒体文件路径。", , SelectFileName)
If SelectFileName = "猪头在想我吗？" Then
SelectFileName = "小朱：" + vbCrLf + "    思夏此刻好想你，好爱你，你呢？" + vbCrLf + "    时间过得很快呀！这一年我们又经历了很多很多。记得上一版本中我说过我希望 Snowman 可以发展起来，透过它你可以随时随地感受到我思念、我无处不在地牵挂。如今 Snowman 下载虽不算火红但也有一定成绩了。然而不知道你是否还能感受到当初的那份感动？我的心就时常为着我们的一切而感动，感动上天赐我一个温柔的你，一个可爱的你。" + vbCrLf + "    你对我的改变我都记在心里，你的付出我也深深感动。在新版本发布之际我希望我们的幸福会变得永恒，希望在你打开 Snowman 的每一刻温馨快乐。"
 MsgBox SelectFileName
Exit Sub
End If


If Len(SelectFileName) > 0 Then
        Lf1.Clear
             LF2.Clear
   
      Lf1.AddItem SelectFileName
              AddFile Lf1.List(Lf1.ListCount - 1)

      Pid = 0
      Lf1.ListIndex = Pid
      LF2.ListIndex = Pid
      LF1_DblClick
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
  Dim m_wCurOptIdx As Long
  Dim txtPath As String
  Dim txtDisplayName As String
    With BI
    .hOwner = Me.hwnd
    nFolder = GetFolderValue(m_wCurOptIdx)
     If SHGetSpecialFolderLocation(ByVal Me.hwnd, ByVal nFolder, IDL) = NOERROR Then
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
                   Lf1.Clear
                    LF2.Clear
                File1.Path = txtPath
  If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AllFiles") = True Then
            Dim result As Long
Dim firstpath As String, dircount As Integer
Dir1.Path = txtPath
firstpath = Dir1.Path
dircount = Dir1.ListCount
result = DirDiver(firstpath, dircount, "")
Pid = 0
      Lf1.ListIndex = Pid
      LF2.ListIndex = Pid
     If Lf1.ListCount > 0 Then
     LF1_DblClick
          End If
   Else
   For i = 0 To File1.ListCount - 1
     Lf1.AddItem File1.Path + "\" + File1.List(i), i
             AddFile Lf1.List(Lf1.ListCount - 1)

    Next
      Pid = 0
      Lf1.ListIndex = Pid
      LF2.ListIndex = Pid
     If Lf1.ListCount > 0 Then
     LF1_DblClick
          End If
     End If
 Me.MousePointer = 0
   End If
   
   
   
   
End Sub

Private Sub a00105_Click()
 Lf1.Clear
 LF2.Clear
 File1.Path = CDRom
          For i = 0 To File1.ListCount - 1
      Lf1.AddItem File1.Path + File1.List(i)
         AddFile Lf1.List(Lf1.ListCount - 1)
            Next
 If Lf1.ListCount > 0 Then
      Pid = 0
      Lf1.ListIndex = Pid
     LF2.ListIndex = Pid
     LF1_DblClick
   End If
End Sub

Private Sub a00106_Click()
On Error Resume Next
Dim Text As String
     Lf1.Clear
                  LF2.Clear
File1.Path = CDRom + "MPEGAV"
 Text = File1.Pattern
 File1.Pattern = "*.dat"
 For i = 0 To File1.ListCount - 1
      Lf1.AddItem File1.Path + "\" + File1.List(i), i
              AddFile Lf1.List(Lf1.ListCount - 1)

Next
File1.Pattern = Text
 If Lf1.ListCount > 0 Then
      Pid = 0
      Lf1.ListIndex = Pid
      LF2.ListIndex = Pid
     LF1_DblClick
   End If
End Sub


Private Sub a00107_Click()
Shell (App.Path + "\SmM_DVD.exe"), vbNormalFocus
End Sub

Private Sub a00201_Click()
If Mper.ShowCaptioning = False Then
Mper.ShowCaptioning = True
a00201.Checked = True
Ig(8).Visible = False
Form_Resize
Else
 Mper.ShowCaptioning = False
a00201.Checked = False
Ig(8).Visible = True
End If
End Sub

Private Sub a00202_Click()
If Len(Mper.Filename) > 0 Then Ly.ShowProp Mper.Filename, Me
If Playing = True Then
 If Len(Str(Track)) = 2 Then Ly.ShowProp CDRom + "Track0" + Right(Str(Track), 1) + ".cda", Me
 If Len(Str(Track)) = 3 Then Ly.ShowProp CDRom + "Track" + Right(Str(Track), 2) + ".cda", Me
End If
End Sub

Private Sub a002020_Click()
Mper.ShowDialog mpShowDialogStatistics
End Sub

Private Sub a00301_Click()
Me.Height = 5145
Me.Width = 7545
End Sub


Private Sub a003012_Click()
On Error Resume Next
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
Mper.DisplaySize = mpFullScreen
mjjhjm.Checked = False
cerde.Checked = False
cbcxbf.Checked = False
a00306.Checked = True
xfg.Checked = False
erg.Checked = True
End Sub

Public Sub RushBm()
On Error Resume Next
Dim ii As Long
Dim nn As String
If Ly.FileExists(myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_X", "")) = True Then
s.Enabled = True
s.Caption = "最后位置 : " + myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_X_I", "")
Else: s.Enabled = False
s.Caption = "退出时的曲目"
End If
For ii = 0 To 4
If ii = 0 Then nn = "A"
If ii = 1 Then nn = "B"
If ii = 2 Then nn = "C"
If ii = 3 Then nn = "D"
If ii = 4 Then nn = "E"
If Ly.FileExists(myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_" + Str(ii), "")) = True Then
a00403(ii).Enabled = True
a00403(ii).Caption = "书签 [" + nn + "] : " + myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_" + Str(ii) + "_I", "")
Else: a00403(ii).Enabled = False

a00403(ii).Caption = "书签 [" + nn + "] : 无内容"
End If
Next
End Sub



Private Sub a00403_Click(Index As Integer)
On Error Resume Next
       Static e As String * 30

       Lf1.Clear
                    LF2.Clear

        Lf1.AddItem myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_" + Str(Index), "")
             AddFile Lf1.List(Lf1.ListCount - 1)
                Lf1.ListIndex = 0
           LF1_DblClick

     
     If UCase(Right(Lf1.List(0), 4)) = ".CDA" Then
                SendMCIString "set cd time format milliseconds", True
               mciSendString "status cd position wait", e, Len(e), True
     
     
        Command = "play cd from " & CStr(Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_" + Str(Index)) - 1000)
 mciSendString Command, 0, 0, 0

SendMCIString "set cd time format tmsf", True

    Else
       Mper.CurrentPosition = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_" + Str(Index)) - 1
   End If
End Sub



Private Sub a00408_Click(Index As Integer)
On Error Resume Next
Static e As String * 30

Dim Text As String
If Len(Mper.Filename) > 0 Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_" + Str(Index), Mper.CurrentPosition
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_" + Str(Index), Mper.Filename
Text = SotPath(Mper.Filename)
If Len(Mper.GetMediaInfoString(mpClipTitle)) > 0 Then Text = Mper.GetMediaInfoString(mpClipTitle)
If Len(Mper.GetMediaInfoString(mpClipAuthor)) > 0 Then Text = Text + " - " + Mper.GetMediaInfoString(mpClipAuthor)
Text = Text + "  " + Gtime(Mper.CurrentPosition) + " / " + Gtime(Mper.Duration)
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_" + Str(Index) + "_I", Text
RushBm
End If
If cd = True Then
          SendMCIString "set cd time format milliseconds", True
               mciSendString "status cd position wait", e, Len(e), 0
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_" + Str(Index), CLng(e)
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_" + Str(Index), Lf1.List(Pid)
Text = LF2.List(Pid)
Text = Text + "  " + Gtime(Int(e / 1000))
    myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_" + Str(Index) + "_I", Text
          SendMCIString "set cd time format tmsf", True

RushBm
End If



End Sub


Sub SetBookMark()
On Error Resume Next
Static e As String * 30

Dim Text As String
If Len(Mper.Filename) > 0 Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_X", Mper.CurrentPosition
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_X", Mper.GetMediaInfoString(mpClipFilename)
Text = SotPath(Mper.Filename)
If Len(Mper.GetMediaInfoString(mpClipTitle)) > 0 Then Text = Mper.GetMediaInfoString(mpClipTitle)
If Len(Mper.GetMediaInfoString(mpClipAuthor)) > 0 Then Text = Text + " - " + Mper.GetMediaInfoString(mpClipAuthor)
Text = Text + "  " + Gtime(Mper.CurrentPosition) + " / " + Gtime(Mper.Duration)
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_X_I", Text
RushBm

Else
  If cd = True Then
          SendMCIString "set cd time format milliseconds", True
               mciSendString "status cd position wait", e, Len(e), 0
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_X", CLng(e)
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_X", Lf1.List(Pid)
Text = LF2.List(Pid)
Text = Text + "  " + Gtime(Int(e / 1000))
    myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_X_I", Text
          SendMCIString "set cd time format tmsf", True

RushBm


Else

myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_X", ""
myWriteINI App.Path + "\SmM_Start.dat", "BookMark", "Bm_X_I", ""
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_X", 0
End If
End If






End Sub

Private Sub a005021_Click()
Dim SelectFileName As String
 SelectFileName = InputBox("请输入万维网地址 (URL) 或指定你要添加的本地媒体文件路径。", , SelectFileName)
  If Len(SelectFileName) > 0 Then
  Lf1.AddItem SelectFileName, Lf1.ListCount
          AddFile Lf1.List(Lf1.ListCount - 1)
End If
End Sub

Private Sub a00503_Click()
If cd = True Then Ig_MouseUp 16, 1, 0, 0, 0
Shell (App.Path + "\SmM_Casket\SmM_Casket.exe"), vbNormalFocus
End Sub

Private Sub a005031_Click()
    Dim BI As BROWSEINFO
  Dim nFolder As Long
  Dim IDL As ITEMIDLIST
  Dim pIdl As Long
  Dim sPath As String
  Dim SHFI As SHFILEINFO
  Dim m_wCurOptIdx As Long
  Dim txtPath As String
  Dim txtDisplayName As String
    With BI
    .hOwner = Me.hwnd
    nFolder = GetFolderValue(m_wCurOptIdx)
     If SHGetSpecialFolderLocation(ByVal Me.hwnd, ByVal nFolder, IDL) = NOERROR Then
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
            Dim result As Long
Dim firstpath As String, dircount As Integer
Dir1.Path = txtPath
firstpath = Dir1.Path
dircount = Dir1.ListCount
result = DirDiver(firstpath, dircount, "")
   Else
   For i = 0 To File1.ListCount - 1
     Lf1.AddItem File1.Path + "\" + File1.List(i), Lf1.ListCount
             AddFile Lf1.List(Lf1.ListCount - 1)

    Next
     End If
    End If
 Me.MousePointer = 0
   End If
End Sub


Private Sub a005063_Click()
If Lf1.ListCount > 0 Then
 Dim Text As String
      CommonDialog1.Filename = ""
        CommonDialog1.DialogTitle = "导出到文件"
           Text = CommonDialog1.Filter
          CommonDialog1.Filter = "列表文件 (*.sml)" & _
          "|*.sml|所有文件 (*.*)|*.*"
             CommonDialog1.ShowSave
          If Len(CommonDialog1.Filename) > 0 Then
            If Right(CommonDialog1.Filename, 4) <> ".sml" Then CommonDialog1.Filename = CommonDialog1.Filename + ".sml"
               If Ly.FileExists(CommonDialog1.Filename) = True Then
                      If MsgBox("文件 """ + CommonDialog1.Filename + """ 已存在,要改写吗?", vbYesNo) = vbYes Then
                                   Open CommonDialog1.Filename For Output As #1
                                    For i = 0 To Lf1.ListCount - 1
                                    Print #1, Lf1.List(i)
                                     Next i
                                    Close (1)
                               Else
                             CommonDialog1.Filter = Text
                          CommonDialog1.DialogTitle = "浏览媒体文件"
           
                               Exit Sub
                               End If
                        Else
                            Open CommonDialog1.Filename For Output As #1
                                    For i = 0 To Lf1.ListCount - 1
                                    Print #1, Lf1.List(i)
                                     Next i
                                    Close (1)
                             End If
                   End If
        End If
                  CommonDialog1.Filter = Text
                          CommonDialog1.DialogTitle = "浏览媒体文件"
    
End Sub

Private Sub a00701_Click()
Shell (App.Path + "\SmM_Settings.exe"), vbNormalFocus
End Sub

Private Sub a00703_Click()
Mper.ShowDialog mpShowDialogOptions
End Sub

Private Sub a00801_Click()
Shell (App.Path + "\SmM_Help.exe")
End Sub

Private Sub a00803_Click()
HttpTo "http://www.gznc.com/h2o/smmud/setup.exe"
End Sub

Private Sub a00805_Click()
HttpTo "http://www.gznc.com/h2o"
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










Private Sub asd2_Click(Index As Integer)
 yj6_Click (Index)
End Sub

Private Sub asdfa_Click(Index As Integer)
asdfawe_Click (Index)
End Sub


Private Sub asdfawe_Click(Index As Integer)
If Index = 0 Then
  For i = 1 To 10
    If bfgb(i).Checked = True Then
    bfgb_Click (i - 1)
    Exit Sub
    End If
  Next
End If
If Index = 1 Then
    For i = 0 To 9
    If bfgb(i).Checked = True Then
    bfgb_Click (i + 1)
    Exit Sub
    End If
  Next
End If
End Sub

Private Sub bf6t5_Click()
Mper.Balance = 9640
xbcty.Checked = False
dgv.Checked = False
xbrth.Checked = False
bf6t5.Checked = True
zxvreg.Checked = False
zxvrs.Checked = True
End Sub

Private Sub bfgb_Click(Index As Integer)
On Error Resume Next
Dim Vol As Long
If bfgb(10).Checked = True Then
  For i = 0 To 10
    If bfgb(i).Checked = True And i <> 10 Then
    Index = i
    Exit For
    End If
  Next
End If
Vol = Int((10 - Index) * 53 / 10)
If Vol > 53 Then Vol = 53
     HB.SetVolume Vol, Vol, 1
     HB.SetVolume Vol, Vol, 2
If Index = 10 Then
      bfgb(10).Checked = True
    rthrtb(10).Checked = True
Exit Sub
End If
For i = 0 To 10
  bfgb(i).Checked = False
  rthrtb(i).Checked = False
Next
    bfgb(Index).Checked = True
    rthrtb(Index).Checked = True
End Sub

Private Sub cbcxbf_Click()
Mper.DisplaySize = mpDefaultSize
mjjhjm.Checked = False
cerde.Checked = False
cbcxbf.Checked = True
a00306.Checked = False
xfg.Checked = False
erg.Checked = False
End Sub

Private Sub cerde_Click()
Mper.DisplaySize = mpHalfSize
mjjhjm.Checked = False
cerde.Checked = True
cbcxbf.Checked = False
a00306.Checked = False
xfg.Checked = False
erg.Checked = False
End Sub





Private Sub OpenMedia(X As Integer, Y As Integer)
On Error Resume Next
        a00105.Enabled = False
         a00106.Enabled = False
         a00108.Enabled = False
         File1.Path = App.Path
          File1.Path = CDRom + "MPEGAV"
         If Ly.FileExists(File1.Path + "\AVSEQ01.DAT") = True Or Ly.FileExists(File1.Path + "\MUSIC01.DAT") Then a00106.Enabled = True
                File1.Path = CDRom
         If Ly.FileExists(File1.Path + "\Track01.cda") = True Then a00105.Enabled = True
         If a00105.Enabled = False And a00106.Enabled = False And File1.Path <> App.Path Then a00108.Enabled = True
           PopupMenu Me.a001, 0, X, Y
End Sub
Private Sub cLT1_ItemClick(ByVal nList As Integer, ByVal nItem As Integer)
On Error Resume Next

Select Case nItem
          Case 1
          Me.Visible = False
          Me.WindowState = 0
          If juyjk.Checked = False Then
          juyjk.Checked = True
          Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LW", Me.Width
          Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LH", Me.Height

          Sf1.SkinPath = Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Path")
          Mper.ShowControls = False
          Mper.ShowStatusBar = False
          Pinfo.Visible = False
          Ig(0).Visible = False
          Ig(1).Visible = False
          Ig(2).Visible = False
          Ig(3).Visible = False
          Ig(4).Visible = False
          Ig(5).Visible = False
          Fm10.Visible = False
          Fm5.Visible = False
          cLT1.Visible = False
          Fm0.Top = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vy")
          Fm0.Left = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vx")
          Fm0.Width = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vw")
          Fm0.Height = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vh")
          Me.Width = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_w")
          Me.Height = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_h")
          LF2.Visible = False
          Mper.Top = 0
          Mper.Left = 0
          Mper.Width = Fm0.Width
          Mper.Height = Fm0.Height
          Ig(7).Visible = False
          Ig(8).Top = 0
          Ig(8).Left = 0
          Ig(8).Picture = LoadPicture(Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Bp"))
              Else
           If Mper.ImageSourceWidth > 0 Then GoTo aa
            For i = 0 To 2
         If sdfewe(i).Enabled = True Then
               Ig(19).Picture = LoadPicture(App.Path + "\SmM_Icos\13.gif")
                If juyjk.Checked = False Then Ig(19).Visible = True
                Exit For
              End If
              Next
            If UCase(Left(Mper.Filename, 2)) = "HT" Or UCase(Left(Mper.Filename, 2)) = "MM" Then
           Ig(19).Picture = LoadPicture(App.Path + "\SmM_Icos\15.gif")
                  If juyjk.Checked = False Then Ig(19).Visible = True

         End If
          If cd = True Then
          Ig(19).Picture = LoadPicture(App.Path + "\SmM_Icos\14.gif")
              Ig(19).Left = Me.Width - 3050 - Ig(19).Width
          End If
aa:
          juyjk.Checked = False
          Sf1.SkinPath = App.Path + "\SmM_Skin"
          Mper.ShowControls = True
          Mper.ShowStatusBar = True
          Pinfo.Visible = True
          Ig(0).Visible = True
          Ig(1).Visible = True
          Ig(2).Visible = True
          Ig(3).Visible = True
          Ig(4).Visible = True
          Ig(5).Visible = True
          Fm10.Visible = True
          Fm5.Visible = True
          cLT1.Visible = True
          Fm0.Top = 0
          Fm0.Left = 0
          LF2.Visible = True
          Mper.Top = 405
          Mper.Left = 45
          Ig(7).Visible = True
          Ig(8).Picture = LoadPicture(App.Path + "\SmM_Skin\Snowman.gif")
          Me.Width = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LW")
          Me.Height = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LH")

          Form_Resize
      End If
          Me.Visible = True
          Case 2
           OpenMedia cLT1.Left + nItemX, cLT1.Top + nItemY
           Case 3
            PopupMenu Me.a002, 0, cLT1.Left + nItemX, cLT1.Top + nItemY
          Case 4
           PopupMenu Me.a003, 0, cLT1.Left + nItemX, cLT1.Top + nItemY
           Case 5
            PopupMenu Me.a004, 0, cLT1.Left + nItemX, cLT1.Top + nItemY
          Case 6
           PopupMenu Me.a005, 0, cLT1.Left + nItemX, cLT1.Top + nItemY
           Case 7
            PopupMenu Me.a006, 0, cLT1.Left + nItemX, cLT1.Top + nItemY
          Case 8
           PopupMenu Me.a007, 0, cLT1.Left + nItemX, cLT1.Top + nItemY
           Case 9
            PopupMenu Me.a008, 0, cLT1.Left + nItemX, cLT1.Top + nItemY
   End Select

End Sub


Private Sub cLT1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 nItemX = X
  nItemY = Y
End Sub

Private Sub cLT1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then If Button = 2 Then PopupMenu Me.b002, 0, X + cLT1.Left, Y + cLT1.Top

End Sub





















Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Timer1.Enabled = False
MoveX = Command1.Left
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If Command1.Left + X - 90 < 46 Then
 Command1.Left = 46
Exit Sub
End If

If Command1.Left + X - 90 > Picture1.Width - 136 Then
Command1.Left = Picture1.Width - 136
Exit Sub
End If

If 46 <= Command1.Left + X - 90 <= Picture1.Width - 136 Then
Command1.Left = Command1.Left + X - 90
Exit Sub
End If


End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static e As String * 30
On Error Resume Next
If Button <> 1 Then Exit Sub
Dim songL As Long
          SendMCIString "set cd time format milliseconds", True
 Command = "status cd length track " & CDid
mciSendString Command, e, Len(e), 0
             songL = e
               mciSendString "status cd position wait", e, Len(e), 0

If (Playing) Then
    Command = "play cd from " & CStr(Int(CLng(e) + songL * (Command1.Left - MoveX) / (Picture1.Width - 230)))
Else
    Command = "seek cd to " & CStr(Int(CLng(e) + songL * (Command1.Left - MoveX) / (Picture1.Width - 230)))
End If
mciSendString Command, 0, 0, 0
SendMCIString "set cd time format tmsf", True
cLT1.SetFocus

Timer1.Enabled = True
End Sub












Private Sub Ct_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
Lf1.Clear
             LF2.Clear

For Each ThisFile In Data.Files

Lf1.AddItem ThisFile
        AddFile Lf1.List(Lf1.ListCount - 1)

Next
Pid = 0
Lf1.ListIndex = Pid
LF2.ListIndex = Pid
LF1_DblClick

End Sub


Private Sub d_Click()
If Ly.FileExists(Ly.GetSysPath + "\sndvol32.exe") = True Then Shell (Ly.GetSysPath + "\sndvol32.exe"), vbNormalFocus
If Ly.FileExists(Ly.GetWinPath + "\sndvol32.exe") = True Then Shell (Ly.GetWinPath + "\sndvol32.exe"), vbNormalFocus
End Sub


Private Sub dfd_Click()
a00501_Click
End Sub

Private Sub dfdfvv_Click(Index As Integer)
sdfewe_Click (Index)
End Sub

Private Sub dfg_Click()
Ig_MouseUp 10, 1, 0, 0, 0
End Sub

Private Sub dfgfg_Click()
Dim SelectFileName As String
 SelectFileName = InputBox("请输入万维网网址 (URL) 或指定你要打开的本地网页文件路径。", , SelectFileName)
If Len(SelectFileName) > 0 Then
HttpTo SelectFileName
     End If

End Sub

Private Sub dfgr_Click()
Ig_MouseUp 17, 1, 0, 0, 0
End Sub


Private Sub dfhgrtht_Click()
Ig_MouseUp 15, 1, 0, 0, 0
End Sub

Private Sub dgv_Click()
xbcty_Click
End Sub

Private Sub Dir1_Change()
    ' Update File listbox to sync with Dir listbox.
    File1.Path = Dir1.Path
End Sub

Private Sub a00108_Click()
On Error Resume Next
Dim result As Long
Me.MousePointer = 11
Dim firstpath As String, dircount As Integer
Dir1.Path = App.Path + "\SmM_DirCheck"
File1.Path = Dir1.Path
Dir1.Path = CDRom
Lf1.Clear
LF2.Clear
firstpath = Dir1.Path
dircount = Dir1.ListCount
result = DirDiver(firstpath, dircount, "")
Me.MousePointer = 0
Pid = 0
  If Lf1.ListCount > 0 Then
      Lf1.ListIndex = Pid
      LF2.ListIndex = Pid
     LF1_DblClick
   End If
End Sub

Private Sub a002045_Click()
If a002045.Checked = False Then
 a002045.Checked = True
 Ig(3).Picture = LoadPicture(App.Path + "\SmM_Icos\11.gif")
 Ig(3).ToolTipText = "随机"
Else
  a002045.Checked = False
   Ig(3).Picture = LoadPicture(App.Path + "\SmM_Icos\10.gif")
  Ig(3).ToolTipText = "原序"
End If
End Sub

Private Sub a00205_Click()
If a00205.Checked = False Then
 a00205.Checked = True
 Ig(4).Picture = LoadPicture(App.Path + "\SmM_Icos\03.gif")
 Ig(4).ToolTipText = "不循环"
Else
  a00205.Checked = False
   Ig(4).Picture = LoadPicture(App.Path + "\SmM_Icos\02.gif")
 Ig(4).ToolTipText = "循环"
End If
End Sub

Private Sub a00501_Click()
CommonDialog1.Filename = ""
          CommonDialog1.ShowOpen
        If Len(CommonDialog1.Filename) > 0 Then
        Lf1.AddItem CommonDialog1.Filename, Lf1.ListCount
                AddFile Lf1.List(Lf1.ListCount - 1)
End If
 End Sub

Private Sub a00101_Click()
CommonDialog1.Filename = ""
          CommonDialog1.ShowOpen
          If Len(CommonDialog1.Filename) > 0 Then
        Lf1.Clear
                     LF2.Clear

      Lf1.AddItem CommonDialog1.Filename
              AddFile Lf1.List(Lf1.ListCount - 1)

      Pid = 0
      Lf1.ListIndex = Pid
      LF2.ListIndex = Pid
      LF1_DblClick
     End If
 End Sub

Private Sub a00505_Click()
Shell (App.Path + "\SmM_Capturer.exe"), vbNormalFocus
End Sub

Private Sub a005074_Click()
Lf1.Clear
             LF2.Clear

End Sub



Private Sub dsfasdsafe_Click()
Ly.ShowProp Lf1.List(LF2.ListIndex), Me
End Sub

Private Sub dsgegd_Click()
Dim SelectFileName As String
Dim Index As Long
If LF2.ListIndex < 0 Then Exit Sub
 SelectFileName = InputBox("请输入该列表项目的新名称。", , LF2.List(LF2.ListIndex))
If Len(SelectFileName) > 0 Then
   Index = LF2.ListIndex
   LF2.RemoveItem Index
   LF2.AddItem SelectFileName, Index
   End If

End Sub

Private Sub erg_Click()
If erg.Checked = False Then
a00306_Click
Else
mjjhjm_Click
End If
End Sub


Private Sub erth_Click()
a00308_Click
End Sub

'Private Sub fdf4_Click()
'Dim SelectFileName As String
' SelectFileName = InputBox("请输入你要拨打的求助电话。", , SelectFileName)
'  If Len(SelectFileName) > 0 Then
'   Ly.Dial SelectFileName, Me
' Else
'   Exit Sub
'End If
'End Sub


Sub SetVw()
On Error Resume Next
If cerde.Checked = True Then mjjhjm_Click
If Mper.ImageSourceWidth = 0 Then Exit Sub
Me.Width = Mper.ImageSourceWidth * 15 + 3100
Me.Height = Mper.ImageSourceHeight * 15 + 2050
End Sub

Private Sub fghdrh_Click()
Shell (App.Path + "\SmM_sysinfo.exe")

End Sub

Private Sub fghgt_Click()
If fghgt.Checked = False Then
   fghgt.Checked = True
   SetVw
Else
   fghgt.Checked = False
End If
End Sub


Private Sub fgrgfrgrgrtg_Click()
Ig_MouseUp 12, 1, 0, 0, 0
End Sub




Private Sub Fm0_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X, Y
End Sub

Private Sub Fm0_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
Lf1.Clear
             LF2.Clear

For Each ThisFile In Data.Files

Lf1.AddItem ThisFile
        AddFile Lf1.List(Lf1.ListCount - 1)

Next
Pid = 0
Lf1.ListIndex = Pid
LF2.ListIndex = Pid
LF1_DblClick
End Sub


Private Sub Fm10_MouseDown(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
If Button = 1 Then
MoveX = X
MoveY = Y
End If
End Sub

Private Sub Fm10_MouseMove(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
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
Lf1.Clear
             LF2.Clear

For Each ThisFile In Data.Files

Lf1.AddItem ThisFile
        AddFile Lf1.List(Lf1.ListCount - 1)

Next
Pid = 0
Lf1.ListIndex = Pid
LF2.ListIndex = Pid
LF1_DblClick
End Sub



Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X, Y
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
Lf1.Clear
             LF2.Clear

For Each ThisFile In Data.Files

Lf1.AddItem ThisFile
        AddFile Lf1.List(Lf1.ListCount - 1)

Next
Pid = 0
Lf1.ListIndex = Pid
LF2.ListIndex = Pid
LF1_DblClick
End Sub



Private Sub Frame1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
Lf1.Clear
             LF2.Clear

For Each ThisFile In Data.Files

Lf1.AddItem ThisFile
        AddFile Lf1.List(Lf1.ListCount - 1)

Next
Pid = 0
Lf1.ListIndex = Pid
LF2.ListIndex = Pid
LF1_DblClick
End Sub


Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button <> 1 Then Exit Sub
Timer1.Enabled = False
MoveX = Command1.Left
Command1.Left = X - 90


End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If X - 90 < 46 Then
 Command1.Left = 46
Exit Sub
End If

If X - 90 > Picture1.Width - 136 Then
Command1.Left = Picture1.Width - 136
Exit Sub
End If

If 46 <= X - 90 <= Picture1.Width - 136 Then
Command1.Left = X - 90
Exit Sub
End If

End Sub

Private Sub Frame2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Static e As String * 30
If Button <> 1 Then Exit Sub

Dim songL As Long
          SendMCIString "set cd time format milliseconds", True
 Command = "status cd length track " & CDid
mciSendString Command, e, Len(e), 0
             songL = e
               mciSendString "status cd position wait", e, Len(e), 0

If (Playing) Then
    Command = "play cd from " & CStr(Int(CLng(e) + songL * (X - 90 - MoveX) / (Picture1.Width - 230)))
Else
    Command = "seek cd to " & CStr(Int(CLng(e) + songL * (X - 90 - MoveX) / (Picture1.Width - 230)))
End If
mciSendString Command, 0, 0, 0
SendMCIString "set cd time format tmsf", True

Timer1.Enabled = True
LF2.SetFocus

End Sub


Private Sub gfdsgawe_Click()
Dim SelectFileName As String
Dim Index As Long
If LF2.ListIndex < 0 Then Exit Sub
 SelectFileName = InputBox("请重新输入该列表项目指向的媒体文件地址。", , Lf1.List(LF2.ListIndex))
If Len(SelectFileName) > 0 Then
   Index = LF2.ListIndex
   Lf1.RemoveItem Index
   Lf1.AddItem SelectFileName, Index
   ReNameB (Index)
   End If

End Sub

Private Sub ghjykjyt_Click()
a00703_Click
End Sub

Private Sub grdfg_Click()
Ig_MouseUp 14, 1, 0, 0, 0
End Sub

Private Sub greg_Click()
Ig_MouseUp 13, 1, 0, 0, 0
End Sub

Private Sub hf6t_Click()
Set Form1 = Nothing
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
Private Sub Sinfo()
On Error Resume Next
Dim Info2 As String
Info2 = ""
If Len(Mper.Filename) > 0 Then
Info2 = "标题   : " + SotPath(Mper.Filename) + vbCrLf
If Len(Mper.GetMediaInfoString(mpClipTitle)) > 0 Then Info2 = "标题   : " + Mper.GetMediaInfoString(mpClipTitle) + vbCrLf

If Len(Mper.GetMediaInfoString(mpClipAuthor)) > 0 Then Info2 = Info2 + "艺术家 : " + Mper.GetMediaInfoString(mpClipAuthor) + vbCrLf
If Len(Mper.GetMediaInfoString(mpClipRating)) = 0 Then
 Info2 = Info2 + "评价   : ☆☆☆" + vbCrLf
 Else
Info2 = Info2 + "评价   : "
For i = 1 To Int(Mper.GetMediaInfoString(mpClipRating))
  Info2 = Info2 + "☆" '
Next
Info2 = Info2 + vbCrLf
End If
If Len(Mper.GetMediaInfoString(mpClipCopyright)) > 0 Then Info2 = Info2 + "版权   : " + Mper.GetMediaInfoString(mpClipCopyright) + vbCrLf
If Len(Mper.GetMediaInfoString(mpClipDescription)) > 0 Then Info2 = Info2 + "描述   : " + Mper.GetMediaInfoString(mpClipDescription) + vbCrLf
Info2 = Info2 + "时间   : " + Gtime(Mper.Duration) + vbCrLf
If Mper.Bandwidth > 0 Then Info2 = Info2 + "质量   :" + Str(Int(Mper.Bandwidth / 1000)) + " 千字节每秒" + vbCrLf
If Mper.ImageSourceWidth > 0 Then
    Info2 = Info2 + "视频   :" + Str(Mper.ImageSourceWidth) + " ×" + Str(Mper.ImageSourceHeight) + vbCrLf
Else
    Info2 = Info2 + "类型   : 音频" + vbCrLf
End If
Info2 = Info2 + "地址   : " + Lf1.List(Pid)
End If

If cd = True Then Info2 = "标题   : " + SotPath(Lf1.List(Pid)) + vbCrLf + "艺术家 : 未知艺术家" + vbCrLf + "唱片集 : 未知唱片集" + vbCrLf + "流派   : 未知流派" + vbCrLf + "时间   : " + Stime + vbCrLf + "地址   : " + Lf1.List(Pid)

If Playing = False And Len(Mper.Filename) = 0 Then
Info2 = "    欢迎使用 Snowman Media 4se 享受无限精彩的数字媒体生活!" + vbCrLf + vbCrLf + "    想了解更多或交流反馈请登陆 流动网络： http://www.gznc.com/h2o 或电邮到 leask@21cn.com。"
End If

If Ly.FileExists(Ly.GetSysPath + "\sndvol32.exe") = True Then
St1.ShowMsg Info2, ICON_INFO, "Snowman Media 4se", 0
Else
MsgBox Info2, vbInformation
End If
End Sub

Private Sub Ig_DblClick(Index As Integer)
If Index = 7 Or Index = 8 Then erg_Click

End Sub

Private Sub Ig_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   MoveX = X
   MoveY = Y
End If

End Sub

Private Sub Ig_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 And (Index = 0 Or Index = 5 Or Index = 6 Or Index = 9 Or (Index = 8 And juyjk.Checked = True)) Then
Me.Left = Me.Left - MoveX + X
Me.Top = Me.Top - MoveY + Y
End If
End Sub

Private Sub Ig_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
Lf1.Clear
             LF2.Clear

For Each ThisFile In Data.Files

Lf1.AddItem ThisFile
        AddFile Lf1.List(Lf1.ListCount - 1)

Next
Pid = 0
Lf1.ListIndex = Pid
LF2.ListIndex = Pid
LF1_DblClick

End Sub

Private Sub il_Click()
Ig_MouseUp 12, 1, 0, 0, 0
End Sub
Private Sub ShowRight(X As Single, Y As Single)



PopupMenu Me.b002, 0, X, Y
End Sub



Private Sub Ig_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim e As String * 40
If Index = 0 Then
If Button = 2 Then ShowRight X, Y
End If

If Index = 2 Then
If Button = 1 Then
SinfoBo = False
Sinfo
End If
If Button = 2 Then ShowRight X + Ig(2).Left, Y + Ig(2).Top
End If

If Index = 3 Then
If Button = 1 Then
a002045_Click
End If
If Button = 2 Then ShowRight X + Ig(3).Left, Y + Ig(3).Top
End If

If Index = 4 Then
If Button = 1 Then
a00205_Click
End If
If Button = 2 Then ShowRight X + Ig(4).Left, Y + Ig(4).Top
End If


If Index = 5 Then
If Button = 2 Then ShowRight X, Y + Ig(5).Top
End If

If Index = 6 Then
If Button = 2 Then ShowRight X + 45, Y + Fm10.Top
End If

If Index = 7 Then
If Button = 2 Then ShowRight X + Ig(7).Left, Y + Ig(7).Top
End If

If Index = 8 Then
If Button = 2 Then ShowRight X + Ig(8).Left, Y + Ig(8).Top
End If

If Index = 9 Then
If Button = 2 Then ShowRight X + Fm5.Left, Y + Fm5.Top
End If


If Index = 10 Then
If Button = 1 Then
If UCase(Left(Mper.Filename, 3)) = "G:\" Or cd = True Then
Ig_MouseUp 16, 1, 0, 0, 0
End If
SendMCIString "Set cd door open", True
End If
If Button = 2 Then ShowRight X + Ig(10).Left, Y + Fm10.Top
End If

If Index = 11 Then
   If Button = 1 Then
      If Lf1.ListCount > 0 Then
        Pid = Pid - 1
  If Pid < 0 Then Pid = Lf1.ListCount - 1
Lf1.ListIndex = Pid
LF2.ListIndex = Pid
LF1_DblClick
End If
End If
If Button = 2 Then ShowRight X + Ig(11).Left, Y + Fm10.Top
End If


If Index = 13 Then
If Button = 1 Then
If cd = True Then
SendMCIString "set cd time format milliseconds", True
mciSendString "status cd position wait", e, Len(e), 0
If (Playing) Then
    Command = "play cd from " & CStr(CLng(e) - 5000)
Else
    Command = "seek cd to " & CStr(CLng(e) - 5000)
End If
mciSendString Command, 0, 0, 0
SendMCIString "set cd time format tmsf", True
'Update
Else
If Len(Mper.Filename) > 0 Then Mper.CurrentPosition = Mper.CurrentPosition + 5
End If
End If
If Button = 2 Then ShowRight X + Ig(13).Left, Y + Fm10.Top
End If


If Index = 14 Then

If Button = 1 Then

If cd = True Then
SendMCIString "play cd", True
Playing = True
Ig(2).Picture = LoadPicture(App.Path + "\SmM_Icos\05.gif")
Exit Sub
End If

If Len(Mper.Filename) > 0 Then
Mper.Play
Exit Sub
End If


If Len(Mper.Filename) = 0 And Playing = False Then
If LF2.ListCount > 0 Then
Pid = 0
LF2.ListIndex = Pid
LF2_DblClick
Else: OpenMedia X + Ig(14).Left, Y + Fm10.Top
End If
End If
End If



If Button = 2 Then ShowRight X + Ig(14).Left, Y + Fm10.Top
End If

If Index = 15 Then
If Button = 1 Then
If cd = True Then
If Playing = True Then
SendMCIString "pause cd", True
'Update
Else
 SendMCIString "play cd", True
'Update


End If
Else
If Len(Mper.Filename) > 0 Then
If Mper.PlayState = mpPlaying Then
Mper.Pause
Else
Mper.Play
End If
End If
End If
End If
If Button = 2 Then ShowRight X + Ig(15).Left, Y + Fm10.Top
End If









If Index = 16 Then
If Button = 1 Then
 Ig(19).Visible = False
If cd = True Then
SendMCIString "stop cd wait", True
'Command = "seek cd to " & 1
'SendMCIString Command, True
Playing = False
Update
  Ct.Caption = "已停止"
  Times.Caption = ""

cd = False
Else
 If Len(Mper.Filename) > 0 Then
 Mper.Filename = "ilxz"

End If
End If
Ig(2).Picture = LoadPicture(App.Path + "\SmM_Icos\08.gif")
End If
If Button = 2 Then ShowRight X + Ig(16).Left, Y + Fm10.Top
End If



If Index = 17 Then
   If Button = 1 Then
        If cd = True Then
          SendMCIString "set cd time format milliseconds", True
               mciSendString "status cd position wait", e, Len(e), 0
              If (Playing) Then
                    Command = "play cd from " & CStr(CLng(e) + 5000)
              Else
                       Command = "seek cd to " & CStr(CLng(e) + 5000)
              End If
mciSendString Command, 0, 0, 0
SendMCIString "set cd time format tmsf", True
'Update
        Else
                   If Len(Mper.Filename) > 0 Then Mper.CurrentPosition = Mper.CurrentPosition + 5
        End If
   End If
If Button = 2 Then ShowRight X + Ig(17).Left, Y + Fm10.Top
End If


If Index = 12 Then
     If Button = 1 Then
         If Lf1.ListCount > 0 Then
    Pid = Pid + 1
           If Pid > Lf1.ListCount - 1 Then Pid = 0
        Lf1.ListIndex = Pid
        LF2.ListIndex = Pid
        LF1_DblClick
          End If
     End If
 If Button = 2 Then ShowRight X + Ig(12).Left, Y + Fm10.Top

End If


If Index = 18 Then
If Button = 1 Then PopupMenu Me.asdf, 0, X + Ig(18).Left, Y + Fm10.Top
If Button = 2 Then ShowRight X + Ig(18).Left, Y + Fm10.Top
End If



If Index = 19 Then
 If Button <> 1 Then Exit Sub
  Dim ConII As Long
 If UCase(Left(Lf1.List(Pid), 2)) = "HT" Or UCase(Left(Lf1.List(Pid), 2)) = "MM" Then
       For i = 1 To Len(Lf1.List(Pid))
             If Mid(Lf1.List(Pid), i, 1) = "/" Then ConII = ConII + 1
             If ConII = 3 Then
                 HttpTo Left(Lf1.List(Pid), i - 1)
                 Exit Sub
             End If
        Next
      Exit Sub
   End If
 For i = 0 To 2
       If sdfewe(i).Enabled = True Then
           sdfewe_Click i
           Exit For
        End If
  Next
If cd = True Then a00503_Click
 
End If
End Sub
Private Sub iuot7oi7to_Click()
a00801_Click
End Sub

Private Sub jhl_Click()
Ig_MouseUp 16, 1, 0, 0, 0
End Sub

Private Sub jkjhl_Click()
Ig_MouseUp 14, 1, 0, 0, 0
End Sub

Private Sub jlk_Click()
Ig_MouseUp 11, 1, 0, 0, 0
End Sub

Private Sub jlkl_Click()
Ig_MouseUp 15, 1, 0, 0, 0
End Sub


Public Sub juyjk_Click()
  cLT1_ItemClick 1, 1
End Sub

Private Sub juyjuk_Click()
a00202_Click
End Sub

Private Sub kuykuk_Click()
a002020_Click
End Sub



Private Sub LF2_DblClick()
If LF2.List(LF2.ListIndex) = "<<== 上一级搜索" Then
a001011_Click
Exit Sub
End If
If Len(Lf1.List(LF2.ListIndex)) = 0 Then
Lf1.RemoveItem (LF2.ListIndex)
LF2.RemoveItem (LF2.ListIndex)

End If

Lf1.ListIndex = LF2.ListIndex
LF1_DblClick
End Sub


Sub LoadSet()
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "MousePu") = True Then
Mper.ClickToPlay = True
Else
Mper.ClickToPlay = False
End If
BSv
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Change", False
End Sub
Sub BSv()
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ScSave") = True Then
       If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "OnlyVideo") = True Then
             If Mper.ImageSourceWidth > 0 Then
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

Private Sub ReLmp()
Mper.SendMouseClickEvents = True
Mper.SendMouseMoveEvents = True
Mper.AnimationAtStart = False
Mper.ClickToPlay = False
Mper.EnableContextMenu = False
Mper.ShowStatusBar = True
Mper.VideoBorder3D = False
Mper.DisplayBackColor = &HFFFFFF
Mper.DisplayForeColor = &HFF0000
End Sub

Private Sub Form_Load()
On Error Resume Next
St1.ShowIcon
If App.PrevInstance = True Then End
Dim Text As String
RushBm
Playing = False
cd = False
Ig(8).Picture = LoadPicture(App.Path + "\SmM_Skin\Snowman.gif")

SendMCIString "open cdaudio alias cd wait shareable", True
SendMCIString "set cd time format tmsf wait", True
SendMCIString "stop cd wait", True

    Open App.Path + "\SmM_Start.dat" For Input As #1
    While Not EOF(1)
    Line Input #1, Text
    Lf1.AddItem RTrim(Text)
    Wend
    Close #1
   Open App.Path + "\SmM_Start.dat" For Output As #1
    For i = 0 To Lf1.ListCount - 1
     Print #1, LeftB(Lf1.List(i), 2000)
    Next i
   Close (1)
 Lf1.Clear

 Open App.Path + "\SmM_List.sml" For Input As #1
    While Not EOF(1)
    Line Input #1, Text
    Lf1.AddItem RTrim(Text)

    Wend
    Close #1
 Open App.Path + "\SmM_NameList.dat" For Input As #1
    While Not EOF(1)
    Line Input #1, Text
    LF2.AddItem RTrim(Text)

    Wend
    Close #1
    
    
    
    Pid = 0
   If Lf1.ListCount > 0 Then
   Lf1.ListIndex = Pid
   LF2.ListIndex = Pid
   End If
Sf1.SkinPath = App.Path + "\SmM_Skin"
   cLT1.AddListImage 1, "S.flake.", 1
    cLT1.AddListImage 1, "打开媒体", 2
   cLT1.AddListImage 1, "当前播放", 3
    cLT1.AddListImage 1, "外观视图", 4
    cLT1.AddListImage 1, "媒体书签", 5
   cLT1.AddListImage 1, "曲目列表", 6
   cLT1.AddListImage 1, "个人媒体", 7
    cLT1.AddListImage 1, "更改选项", 8
   cLT1.AddListImage 1, "帮助支持", 9
  cd = False
Ig(2).Picture = LoadPicture(App.Path + "\SmM_Icos\08.gif")
Ig(3).Picture = LoadPicture(App.Path + "\SmM_Icos\10.gif")
Ig(4).Picture = LoadPicture(App.Path + "\SmM_Icos\02.gif")
          CommonDialog1.Filter = "媒体文件 (*.smr;*.smm;*.sma;*.smv;*.sml;*.ilxz;*.asf;*.asx;*.wm;*.wmx;*.wmp;*.wma;*.wax;*.wmv;*.wvx;*.vob;*.cda;*.wav;*.avi;*.mpeg;*.mpg;*.mpe;*.m1v;*.mp;*.mpv2;*.mpv;*.mpa;*.mp3;*.m3u;*.mid;*.midi;*.rmi;*.ivf;*.aif;*.aifc;*.aiff;*.au;*.snd;*.swf)|*.smr;*.smm;*.sma;*.smv;*.sml;*.ilxz;*.asf;*.asx;*.wm;*.wmx;*.wmp;*.wma;*.wax;*.wmv;*.wvx;*.vob;*.cda;*.wav;*.avi;*.mpeg;*.mpg;*.mpe;*.m1v;*.mp;*.mpv2;*.mpv;*.mpa;*.mp3;*.m3u;*.mid;*.midi;*.rmi;*.ivf;*.aif;*.aifc;*.aiff;*.au;*.snd;*.swf|" & _
          "Real 媒体文件 (*.ra;*.rm;*.rmm;*.r1m;*.rom;*.mns;*.rp;*.rtx;*.rt;*.ram;*.rmx;*.rmj;*.rms;*.pls;*.xpl;*.smi;*.smil;*.mnd;*.rmvb;*.ssm;*.rv;*.sdp;*.r3t;*.acp;*.la1;*.lar;*.vpg)|*.ra;*.rm;*.rmm;*.r1m;*.rom;*.mns;*.rp;*.rtx;*.rt;*.ram;*.rmx;*.rmj;*.rms;*.pls;*.xpl;*.smi;*.smil;*.mnd;*.rmvb;*.ssm;*.rv;*.sdp;*.r3t;*.acp;*.la1;*.lar;*.vpg|VCD;DVD 视频 (*.dat;*.vob)|*.dat;*.vob|图片文件 (*.bmp;*jpg;*.gif;*.dib;*.emf;*.wmf)|*.bmp;*jpg;*.gif;*.dib;*.emf;*.wmf|所有文件 (*.*)|*.*"
          CommonDialog1.FilterIndex = 1
File1.Pattern = "*.smr;*.smm;*.sma;*.smv;*.sml;*.ilxz;*.asf;*.asx;*.wm;*.wmx;*.wmp;*.wma;*.wax;*.wmv;*.wvx;*.vob;*.cda;*.wav;*.avi;*.mpeg;*.mpg;*.mpe;*.m1v;*.mp;*.mpv2;*.mpv;*.mpa;*.mp3;*.m3u;*.mid;*.midi;*.rmi;*.ivf;*.aif;*.aifc;*.aiff;*.au;*.snd;*.swf"
  CDRom = Left(Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "CDRom"), 3)
    Dim l As Long
    Dim wHotkey As Long
    wHotkey = (HOTKEYF_ALT Or HOTKEYF_CONTROL) * 256 + vbKeyS
    l = SendMessage(Me.hwnd, WM_SETHOTKEY, wHotkey, 0)






If Ly.FileExists(Ly.GetSysPath + "\sndvol32.exe") = False Then


xtgf4.Enabled = False
gmh.Enabled = False



End If


Me.Top = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LT")
Me.Left = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LL")
 Me.Width = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LW")
 Me.Height = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LH")


If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Mt") = 1 Then a00308_Click
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

If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Skin") = 1 Then juyjk_Click
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Ftv") = 1 Then fghgt_Click

If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Sf") = 1 Then a003012_Click

LoadSet
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoMedia") = True Then Shell (App.Path + "\SmM_Types.exe")
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "PlayFile") <> True Then
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoStart") = True Then Ly.PlayWav App.Path + "\SmM_Medias\Wellcome.wav", False
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoCnt") = True Then s_Click


End If

Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Sting", True
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "RunTime") = False Then
Ly.CenterForm Me
Form2.Show
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "RunTime", True
End If
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Mt") = 1 Then a00308_Click
End Sub



Private Sub LF2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Len(Lf1.List(LF2.ListIndex)) = 0 And LF2.List(LF2.ListIndex) <> "<<== 上一级搜索" Then
Lf1.RemoveItem (LF2.ListIndex)
LF2.RemoveItem (LF2.ListIndex)

End If

If Button <> 1 Then Exit Sub
ReAdd = LF2.ListIndex
End Sub

Private Sub LF2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim lind  As Long
Dim tem As Long
Dim lXPoint As Long
Dim lYPoint As Long
If Y = 0 Then Exit Sub

lXPoint = CLng(X / Screen.TwipsPerPixelX)
lYPoint = CLng(Y / Screen.TwipsPerPixelY)
lind = SendMessage(LF2.hwnd, LB_ITEMFROMPOINT, 0, ByVal ((lYPoint * 65536) + lXPoint))
Oline = lind
If lind <= LF2.ListCount Then
LF2.ListIndex = lind
LF2.ToolTipText = "[" + Str(lind + 1) + " -" + Str(LF2.ListCount) + " ]  " + LF2.List(lind)
Else
LF2.ToolTipText = ""
End If
If LF2.ListCount < 1 Then Exit Sub
If Button <> 1 Then
Picture2.Visible = False
Exit Sub
End If
Picture2.Visible = True
tem = Int(Y / TextHeight("x"))
If tem < 0 Then tem = 0
If tem > LF2.Height / TextHeight("x") Then tem = Int(LF2.Height / TextHeight("x"))
If tem > LF2.ListCount Then tem = LF2.ListCount
Picture2.Top = tem * TextHeight("x") + LF2.Top - 7



End Sub

Private Sub LF2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Picture2.Visible = False
Dim AbbW As Boolean
If Button = 2 Then
PopupMenu Me.a005, 0, LF2.Left + X, LF2.Top + Y
End If


If Button <> 1 Then Exit Sub
If ReAdd = LF2.ListIndex Then Exit Sub

If (Picture2.Top <> LF2.Top - 7 And Oline > LF2.ListCount) = True Or Picture2.Top >= LF2.Top - 7 + Int(LF2.Height / TextHeight("x")) * TextHeight("x") Then
If ReAdd < 0 Then Exit Sub
Lf1.AddItem Lf1.List(ReAdd)
LF2.AddItem LF2.List(ReAdd)
Lf1.RemoveItem ReAdd
LF2.RemoveItem ReAdd
Exit Sub
End If

If -1 < ReAdd < LF2.ListCount Then
If LF2.ListIndex < ReAdd Then
AbbW = True
Else
AbbW = False
End If
Lf1.AddItem Lf1.List(ReAdd), LF2.ListIndex
LF2.AddItem LF2.List(ReAdd), LF2.ListIndex
If AbbW = True Then
Lf1.RemoveItem ReAdd + 1
LF2.RemoveItem ReAdd + 1
Else
Lf1.RemoveItem ReAdd
LF2.RemoveItem ReAdd
End If
End If


End Sub


Private Sub LF2_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
 LF2_MouseMove 1, 0, X, Y
End Sub

Private Sub mjjhjm_Click()
Mper.DisplaySize = mpFitToSize
mjjhjm.Checked = True
cerde.Checked = False
cbcxbf.Checked = False
a00306.Checked = False
xfg.Checked = False
erg.Checked = False
End Sub

Private Sub ReName()
 LF2.Clear
For i = 0 To Lf1.ListCount - 1
   AddFile Lf1.List(i)
Next

End Sub
Public Sub ReNameB(PiN As Long)
Dim Text As String
   Dim File As String
 
    File = Lf1.List(PiN)
       Dim Jd As Long, File2 As String
       If Len(File) = 0 Then Exit Sub

   Me.MousePointer = 11
    LF2.RemoveItem PiN

     File = Lf1.List(PiN)
     If Len(Mper.Filename) > 0 Then
         File2 = Mper.Filename
      Jd = Mper.CurrentPosition
      Else: File2 = "None"
      End If
      Mper.Filename = File
      If (UCase(Right(File, 4)) = ".DAT" And UCase(Left(File, 3)) = CDRom) Or UCase(Right(File, 4)) = ".CDA" Then
         Text = SotPath(File) + " - 未知艺术家"
      GoTo bb
       End If
     Text = SotPath(File)
    If Len(Mper.GetMediaInfoString(mpClipTitle)) > 0 Then Text = Mper.GetMediaInfoString(mpClipTitle)
    If Len(Mper.GetMediaInfoString(mpClipAuthor)) > 0 Then Text = Text + " - " + Mper.GetMediaInfoString(mpClipAuthor)

   
bb:
    Mper.Filename = File2
       Mper.CurrentPosition = Jd

  
  LF2.AddItem Text, PiN
File = ""
Me.MousePointer = 0


End Sub





Private Sub Mper_MouseUp(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.b002, 0, X + Mper.Left + Fm0.Left, Y + Mper.Top + Fm0.Top

End Sub

Private Sub Mper_NewStream()
On Error Resume Next
If Playing = True Then
SendMCIString "stop cd wait", True
Playing = False
cd = False
End If
If UCase(Left(Mper.Filename, 2)) = "HT" Or UCase(Left(Mper.Filename, 2)) = "MM" Then
      Ig(2).Picture = LoadPicture(App.Path + "\SmM_Icos\09.gif")
Else
If Mper.ImageSourceHeight = 0 Then
     Ig(2).Picture = LoadPicture(App.Path + "\SmM_Icos\01.gif")
  Else
     If fghgt.Checked = True Then SetVw
    
         If UCase(Right(Mper.Filename, 4)) = ".BMP" Or UCase(Right(Mper.Filename, 4)) = ".JPG" Or UCase(Right(Mper.Filename, 4)) = ".GIF" Or UCase(Right(Mper.Filename, 4)) = ".WMF" Or UCase(Right(Mper.Filename, 4)) = ".EMF" Or UCase(Right(Mper.Filename, 4)) = ".DIB" Then
               Ig(2).Picture = LoadPicture(App.Path + "\SmM_Icos\04.gif")
          Else
                If UCase(Left(Mper.Filename, 3)) = CDRom And UCase(Right(Mper.Filename, 4)) = ".DAT" Then
                    Ig(2).Picture = LoadPicture(App.Path + "\SmM_Icos\12.gif")
                        Else
                        Ig(2).Picture = LoadPicture(App.Path + "\SmM_Icos\06.gif")
                  End If
           End If
End If
End If
BSv
sdfewe(0).Enabled = True
dfdfvv(0).Enabled = True
Info = "[" + Str(Pid + 1) + " -" + Str(Lf1.ListCount) + " ]  " + "标题:" + SotPath(Mper.Filename) + "  "
If Len(Mper.GetMediaInfoString(mpClipTitle)) > 0 Then Info = "[" + Str(Pid + 1) + " -" + Str(Lf1.ListCount) + " ]  " + "标题:" + Mper.GetMediaInfoString(mpClipTitle) + "  "
If Len(Mper.GetMediaInfoString(mpClipAuthor)) > 0 Then
Info = Info + "艺术家:" + Mper.GetMediaInfoString(mpClipAuthor) + "  "
sdfewe(1).Enabled = True
dfdfvv(1).Enabled = True
End If


If Len(Mper.GetMediaInfoString(mpClipRating)) = 0 Then
 Info = Info + "评价:☆☆☆  "
 Else
Info = Info + "评价:"
For i = 1 To Int(Mper.GetMediaInfoString(mpClipRating))
  Info = Info + "☆" '
Next
Info = Info + "  "
End If


If Len(Mper.GetMediaInfoString(mpClipCopyright)) > 0 Then Info = Info + "版权:" + Mper.GetMediaInfoString(mpClipCopyright) + "  "
If Len(Mper.GetMediaInfoString(mpClipDescription)) > 0 Then
 Info = Info + "描述:" + Mper.GetMediaInfoString(mpClipDescription) + "  "
sdfewe(2).Enabled = True
dfdfvv(2).Enabled = True

End If
Info = Info + "时间: " + Gtime(Mper.Duration) + "  "
If Mper.Bandwidth > 0 Then Info = Info + "质量:" + Str(Int(Mper.Bandwidth / 1000)) + " 千字节每秒  "
If Mper.ImageSourceWidth > 0 Then
    Info = Info + "视频:" + Str(Mper.ImageSourceWidth) + " ×" + Str(Mper.ImageSourceHeight) + "  " '+ " @" +left( Ly.GetDisplay,4)+""+right(ly.GetDisplay,3)) + "  "
Else
    Ig(19).Visible = True

    Info = Info + "类型:音频  "
End If
Info = Info + "地址:" + Lf1.List(Pid)
For i = 0 To 2
 If Mper.ImageSourceWidth > 0 Then Exit Sub
If sdfewe(i).Enabled = True Then
   Ig(19).Picture = LoadPicture(App.Path + "\SmM_Icos\13.gif")
   If juyjk.Checked = False Then Ig(19).Visible = True
   Exit For
End If
Next
 If UCase(Left(Mper.Filename, 2)) = "HT" Or UCase(Left(Mper.Filename, 2)) = "MM" Then
 Ig(19).Picture = LoadPicture(App.Path + "\SmM_Icos\15.gif")
    If juyjk.Checked = False Then Ig(19).Visible = True

 End If
'End If

 Ig(19).Left = Me.Width - 3100 - Ig(19).Width
 'Tinfo.Left = 100

If Ly.FileExists(Ly.GetSysPath + "\sndvol32.exe") = True Then
SinfoBo = True
Sinfo
End If

End Sub




Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button <> 1 Then Exit Sub
Timer1.Enabled = False
MoveX = Command1.Left
Command1.Left = X

End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
If X < 46 Then
 Command1.Left = 46
Exit Sub
End If

If X > Picture1.Width - 136 Then
Command1.Left = Picture1.Width - 136
Exit Sub
End If

If 46 <= X <= Picture1.Width - 136 Then
Command1.Left = X
Exit Sub
End If

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static e As String * 30
On Error Resume Next

If Button <> 1 Then Exit Sub

Dim songL As Long
          SendMCIString "set cd time format milliseconds", True
 Command = "status cd length track " & CDid
mciSendString Command, e, Len(e), 0
             songL = e
               mciSendString "status cd position wait", e, Len(e), 0

If (Playing) Then
    Command = "play cd from " & CStr(Int(CLng(e) + songL * (X - MoveX) / (Picture1.Width - 230)))
Else
    Command = "seek cd to " & CStr(Int(CLng(e) + songL * (X - MoveX) / (Picture1.Width - 230)))
End If
mciSendString Command, 0, 0, 0
SendMCIString "set cd time format tmsf", True

Timer1.Enabled = True
LF2.SetFocus

End Sub


Private Sub PInfo_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
Lf1.Clear
             LF2.Clear

For Each ThisFile In Data.Files

Lf1.AddItem ThisFile
        AddFile Lf1.List(Lf1.ListCount - 1)

Next
Pid = 0
Lf1.ListIndex = Pid
LF2.ListIndex = Pid
LF1_DblClick


End Sub


Private Sub rg_Click()
a00701_Click
End Sub



Private Sub rthrtb_Click(Index As Integer)
bfgb_Click (Index)
End Sub

Private Sub s_Click()
 On Error Resume Next
       Static e As String * 30
    If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_X") = 0 Then Exit Sub
          Pid = -1
     For i = 0 To Lf1.ListCount - 1
      If Lf1.List(i) = Left(myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_X", ""), Len(Lf1.List(i))) Then
      Pid = i
      LF2.ListIndex = i
      LF2_DblClick
      Exit For
      End If
      Next
      If Pid = -1 Then
      Lf1.Clear
          LF2.Clear
         Pid = 0
        Lf1.AddItem myReadINI(App.Path + "\SmM_Start.dat", "BookMark", "Bm_X", "")
             AddFile Lf1.List(0)
                  LF2.ListIndex = Pid
                  LF2_DblClick
     End If
  If UCase(Right(Lf1.List(0), 4)) = ".CDA" Then
          SendMCIString "set cd time format milliseconds", True
               mciSendString "status cd position wait", e, Len(e), True
     
     
        Command = "play cd from " & CStr(Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_X") - 1000)
 mciSendString Command, 0, 0, 0

SendMCIString "set cd time format tmsf", True

     Else
      Mper.CurrentPosition = Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "BookMark_X") - 1
    End If
End Sub

Private Sub sdf_Click()
a00101_Click
End Sub



Private Sub sdfewe_Click(Index As Integer)
Dim Tfor As String
On Error Resume Next
If Len(Mper.Filename) = 0 Then Exit Sub
If Index = 0 Then
Tfor = "http://www.gracenote.com/php/search2.php3?f=track&q="
Tfor = Tfor + SotPath(Mper.Filename)
If Len(Mper.GetMediaInfoString(mpClipTitle)) > 0 Then Tfor = "http://www.gracenote.com/php/search2.php3?f=track&q=" + Mper.GetMediaInfoString(mpClipTitle)

End If

If Index = 1 Then Tfor = "http://www.gracenote.com/php/search2.php3?f=artist&q=" + Mper.GetMediaInfoString(mpClipAuthor)

If Index = 3 Then Tfor = "http://www.gracenote.com/php/search2.php3?f=disc&q=" + Mper.GetMediaInfoString(mpClipDescription)


HttpTo Tfor


End Sub





Private Sub HttpTo(www As String)
Ly.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "NetShow", True
Ly.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "NetFile", www
Shell App.Path + "\SmM_IntBrowser.exe", vbMinimizedFocus
End Sub







Private Sub sdfewer_Click()
HttpTo "http://www.gracenote.com"
End Sub

Private Sub sdfg_Click()
cLT1_ItemClick 1, 1
End Sub




Private Sub SF1_OnSkinNotify(ByVal SkinClass As String, ByVal SkinEvent As String)
    Select Case SkinClass
    Case "A"
         Ig_MouseUp 14, 1, 0, 0, 0
    Case "B"
         Ig_MouseUp 15, 1, 0, 0, 0
    Case "C"
         Ig_MouseUp 11, 1, 0, 0, 0
    Case "D"
         Ig_MouseUp 12, 1, 0, 0, 0
    Case "E"
         Ig_MouseUp 13, 1, 0, 0, 0
    Case "F"
         Ig_MouseUp 16, 1, 0, 0, 0
    Case "G"
         Ig_MouseUp 10, 1, 0, 0, 0
    Case "H"
         Ig_MouseUp 17, 1, 0, 0, 0
    Case "I"
         Ig_MouseUp 1, 1, 0, 0, 0
    Case "J"
         a00101_Click
    Case "K"
         a00501_Click
    Case "L"
         erg_Click
    Case "M"
        a00811_Click
    Case "N"
        a00308_Click
    Case "O"
        a00701_Click
    Case "P"
        a00801_Click
    Case "Q"
        Ig_MouseUp 2, 1, 0, 0, 0
    Case "R"
         Ig_MouseUp 3, 1, 0, 0, 0
    Case "S"
         Ig_MouseUp 4, 1, 0, 0, 0
    Case "T"
         Ig_MouseUp 5, 1, 0, 0, 0
    Case "U"
         Ig_MouseUp 6, 1, 0, 0, 0
    Case "V"
         Ig_MouseUp 7, 1, 0, 0, 0
    Case "W"
         Ig_MouseUp 8, 1, 0, 0, 0
    Case "X"
         Ig_MouseUp 9, 1, 0, 0, 0
    Case "Y"
         Ig_MouseUp 10, 1, 0, 0, 0
    Case "Z"
         Ig_MouseUp 18, 1, 0, 0, 0
    Case "Aa"
         Ig_MouseUp 19, 1, 0, 0, 0
    Case "min"
         Me.WindowState = 1
    Case "all"
          cLT1_ItemClick 1, 1
    End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next
If juyjk.Checked = True Then Exit Sub
If Me.WindowState <> 1 Then
If Me.Width < 7605 Then Me.Width = 7605
If Me.Height < 5205 Then Me.Height = 5205

Fm0.Visible = False
Fm0.Width = Me.Width
Fm0.Height = Me.Height
Ig(0).Width = Me.Width
Ig(1).Width = Me.Width
Ig(5).Width = Me.Width
Fm10.Width = Me.Width
'mper.Width = Me.Width - 3180
Mper.Width = Me.Width - 3100
Ig(19).Left = Me.Width - 3050 - Ig(19).Width

LF2.Left = Me.Width - 3010
Picture2.Left = Me.Width - 2990

'Ig10.Left = Me.Width - 2475
Ig(2).Left = Me.Width - 2205
Ig(3).Left = Me.Width - 1525
Ig(4).Left = Me.Width - 845
Frame1.Top = Me.Height - 825
Frame1.Width = Me.Width - 3250
'TrackTime.Left = Frame1.Width - 2535
Times.Left = Frame1.Width - 1550
cLT1.Left = Me.Width - 3060

Fm5.Left = Me.Width - 3110
LF2.Height = Me.Height - 2115
Mper.Height = Me.Height - 915
Ig(7).Width = Me.Width - 3150
Ig(7).Height = Me.Height - 2550
'PInfo.Width = Me.Width - 2500
Pinfo.Width = Me.Width - 2450
Ig(5).Top = Me.Height - 1875
Ig(19).Top = Me.Height - 2150

Ig(1).Top = Me.Height - 1560
Fm10.Top = Me.Height - 1290
'Ig3.Top = Me.Height - 1290
cLT1.Top = Me.Height - 1530
Fm5.Top = Me.Height - 1530
Ig(8).Left = (Me.Width - 3135 + Ig(8).Width) / 2 - Ig(8).Width + 45
Ig(8).Top = (Me.Height - 2010 + Ig(8).Height) / 2 - Ig(8).Height + 275
Fm0.Visible = True
Frame2.Width = Me.Width - 3200
Frame2.Top = Me.Height - 1525

Picture1.Width = Frame2.Width - 100

'Frame2.Visible = False
'If yj6(0).Checked = True Then Form_Resize
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
Ig_MouseUp 16, 1, 0, 0, 0
 Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "RealPlay", False

If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Clean") = True Then
Lf1.Clear
             LF2.Clear
End If

Open App.Path + "\SmM_List.sml" For Output As #1
    For i = 0 To Lf1.ListCount - 1
     Print #1, Lf1.List(i)
    Next i
   Close (1)
Open App.Path + "\SmM_NameList.dat" For Output As #1
    For i = 0 To LF2.ListCount - 1
     Print #1, LF2.List(i)
    Next i
   Close (1)

Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Sting", False
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "StartUp") = True Then Shell (App.Path + "\SmM_Helper.exe")
 Dim xx As Long
 Dim yy As Long

If Me.WindowState = 2 Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LM", 1
Else
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LM", 0
End If
Me.WindowState = 0
 Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LT", Me.Top
 Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LL", Me.Left
If a00308.Checked = True Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Mt", 1
Else
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Mt", 0
End If
If a003012.Checked = True Then
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Sf", 1
Else
Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_Sf", 0
 Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LW", Me.Width
 Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SmM_LH", Me.Height

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
rc = mciSendString(Cmd, 0, 0, hwnd)
If (fShowError And rc <> 0) Then
    mciGetErrorString rc, errStr, Len(errStr)
    'MsgBox errStr
End If
SendMCIString = (rc = 0)
End Function
Private Sub GoNext()
On Error Resume Next
 Ig(19).Visible = False

If a002045.Checked = False Then
   If Pid >= Lf1.ListCount - 1 Then
    Pid = 0
         Lf1.ListIndex = Pid
         LF2.ListIndex = Pid
         
        If a00205.Checked = False Then
      Mper.Filename = "ilxz"
 
          Else
          Lf1.ListIndex = Pid
          LF2.ListIndex = Pid
          
              LF1_DblClick
           End If
    Else
Lf1.ListIndex = Pid + 1
LF2.ListIndex = Pid + 1

LF1_DblClick
 End If
Else
i = Int((Lf1.ListCount - 1) * Rnd)
Pid = i
Lf1.ListIndex = Pid
LF2.ListIndex = Pid

LF1_DblClick
End If

End Sub
Function SotPath(T$) As String
On Error Resume Next
Dim ii As Long
'Dim TeFo As String
  '  Dir1.Path = App.Path
  '    Dir1.Path = T$
  '  If Dir1.Path <> App.Path Then
  '        For ii = 1 To Len(T$)
   ''           If Mid(T$, ii, 1) = "\" Then TeFo = Right(T$, Len(T$) - ii) + "  < 文件夹 >"
   '       Next
   '    SotPath = TeFo
   '    Exit Function
   '  End If
   If UCase(Right(T$, 4)) = ".CDA" Then
       SotPath = "CD 曲目" + Str(Int(Right(Left(T$, 10), 2)))
         Exit Function
  End If
      
      If UCase(Right(T$, 4)) = ".DAT" And UCase(Left(T$, 3)) = CDRom Then
      SotPath = "VCD 曲目" + Str(Int(Right(Left(T$, 17), 2)))
   Exit Function
   End If




Dim X%, Ct%
SotPath$ = T$
X% = InStr(T$, "\")
Do While X%
Ct% = X%
X% = InStr(Ct% + 1, T$, "\")
Loop
If Ct% > 0 Then SotPath$ = Mid$(T$, Ct% + 1)
For ii = 1 To Len(SotPath) - 1

If Left(Right(SotPath, ii), 1) = "." Then
SotPath = Left(SotPath, Len(SotPath) - ii)
Exit Function
End If
Next
End Function
Public Sub AddFile(File As String)
On Error Resume Next
    Dim Text As String
           Dim Jd As Long, File2 As String
     If Len(File) = 0 Then Exit Sub
        Me.MousePointer = 11
    If Len(Mper.Filename) > 0 Then
         File2 = Mper.Filename
      Jd = Mper.CurrentPosition
      Else: File2 = "None"
      End If
      If (UCase(Right(File, 4)) = ".DAT" And UCase(Left(File, 3)) = CDRom) Or UCase(Right(File, 4)) = ".CDA" Then
         Text = SotPath(File) + " - 未知艺术家"
      GoTo bb
       End If
     Text = SotPath(File)
    If Len(Mper.GetMediaInfoString(mpClipTitle)) > 0 Then Text = Mper.GetMediaInfoString(mpClipTitle)
    If Len(Mper.GetMediaInfoString(mpClipAuthor)) > 0 Then Text = Text + " - " + Mper.GetMediaInfoString(mpClipAuthor)
bb:
     Mper.Filename = File2
       Mper.CurrentPosition = Jd
 LF2.AddItem Text
File = ""
Me.MousePointer = 0
End Sub


Private Sub Lf2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
 Dim tem As Long
  tem = Int(Y / TextHeight("I"))
If tem < 0 Then tem = 0
If tem > LF2.Height / TextHeight("I") Then tem = LF2.Height / TextHeight("I")
If tem > LF2.ListCount Then tem = LF2.ListCount
     Dim ThisFile As Variant
    For Each ThisFile In Data.Files
        Lf1.AddItem ThisFile, tem
           LF2.AddItem ThisFile, tem
            ReNameB tem
         'tem = tem + 1
    Next
End Sub


Public Sub LF1_DblClick()
On Error Resume Next
Dim Text As String
Static e As String * 30
Dim Filename As String
Me.MousePointer = 11
    Dim Rf(27) As String
     Rf(0) = "PLS"
     Rf(1) = "XPL"
     Rf(2) = "SMI"
     Rf(3) = "MND"
     Rf(4) = "MIL"
     Rf(5) = "MVB"
     Rf(6) = "SSM"
     Rf(7) = ".RV"
     Rf(8) = "R3T"
     Rf(9) = "SDP"
     Rf(10) = "ACP"
     Rf(11) = "AVS"
     Rf(12) = "LA1"
     Rf(13) = "LAR"
     Rf(14) = "VPG"
     Rf(15) = ".RM"
     Rf(16) = ".RA"
     Rf(17) = "RMM"
     Rf(18) = "R1M"
     Rf(19) = "ROM"
     Rf(20) = "MNS"
     Rf(21) = ".RP"
     Rf(22) = "RTX"
     Rf(23) = ".RT"
     Rf(24) = "RAM"
     Rf(25) = "RMX"
     Rf(26) = "RMJ"
     Rf(27) = "RMS"
  
  
    Pid = Lf1.ListIndex
    Dir1.Path = App.Path
      Dir1.Path = Lf1.List(Pid)
     
      
      If Dir1.Path <> App.Path Then
   If Dir1.ListCount <> 0 Or File1.ListCount <> 0 Then
  If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AllFiles") = True Then
            Dim result As Long
Dim firstpath As String, dircount As Integer
dircount = Dir1.ListCount
result = DirDiver(firstpath, dircount, "")
    Lf1.RemoveItem (Pid)
      LF2.RemoveItem (Pid)
      Lf1.ListIndex = Pid
      LF2.ListIndex = Pid
If Lf1.ListCount > 0 Then
   LF1_DblClick
 End If
   Else
   For i = 0 To File1.ListCount - 1
     Lf1.AddItem File1.Path + "\" + File1.List(i), i
    Next
    Lf1.RemoveItem (Pid)
          LF2.RemoveItem (Pid)

     Lf1.ListIndex = Pid
     LF2.ListIndex = Pid
     If Lf1.ListCount > 0 Then
     LF1_DblClick
     ReName
    End If
     

     End If
    End If
 
 Me.MousePointer = 0
 Exit Sub
 End If
  
  If UCase(Right(Lf1.List(Pid), 4)) = ".SMR" Then
  Filename = Lf1.List(Pid)
   Lf1.Clear
  Open Filename For Input As #1
    While Not EOF(1)
    Line Input #1, Text
    Lf1.AddItem RTrim(Text)

    Wend
    Close #1
      LF2.Clear

    Open Left(Filename, Len(Filename) - 3) + "dat" For Input As #1
    While Not EOF(1)
    Line Input #1, Text
    LF2.AddItem RTrim(Text)

    Wend
    Close #1
   LF2.AddItem "<<== 上一级搜索"
     

Me.MousePointer = 0
Exit Sub
 End If

 If UCase(Right(Lf1.List(Pid), 4)) = ".SML" Then
   Open Lf1.List(Pid) For Input As #1
    While Not EOF(1)
    Line Input #1, Text
    Lf1.AddItem RTrim(Text)

    Wend
    Close #1
 Lf1.RemoveItem (Pid)
       LF2.RemoveItem (Pid)
       Lf1.ListIndex = Pid

       LF2.ListIndex = Pid
     If Lf1.ListCount > 0 Then
LF1_DblClick
ReName
 End If

Me.MousePointer = 0
Exit Sub
 End If
 
 
 
 If UCase(Right(Lf1.List(Pid), 4)) = ".SWF" Then
 Mper.Filename = "ilxz"
 Shell (App.Path + "\SmM_Flash.exe " + Lf1.List(Pid)), vbNormalFocus
 Me.MousePointer = 0
 Exit Sub
 End If
  
 For i = 0 To 27
   If UCase(Right(Lf1.List(Pid), 3)) = Rf(i) Then
   Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "RealPlay", True
  Mper.Filename = "ilxz"

   myWriteINI Ly.GetSysPath + "\SmM_RealMedia.dat", "Real", "Filename", Lf1.List(Pid)
  If Ly.FileExists(Ly.GetSysPath + "\rmoc3260.dll") = True Then Shell Ly.GetSysPath + "\SmM_RealMedia.exe ", vbNormalFocus
 Me.MousePointer = 0
 Exit Sub
 End If
 
 Next
  
  
  If UCase(Right(Lf1.List(Pid), 4)) = ".CDA" And Ly.FileExists(Lf1.List(Pid)) = True Then
             Ig(19).Picture = LoadPicture(App.Path + "\SmM_Icos\14.gif")
              Ig(19).Left = Me.Width - 3050 - Ig(19).Width

             If juyjk.Checked = False Then Ig(19).Visible = True
              SendMCIString "open cdaudio alias cd wait shareable", True
             SendMCIString "set cd time format tmsf wait", True
                     '  mciSendString "status cd number of tracks wait", e, Len(e), 0
     
                    ' Update
                   Playing = True
                Ig(2).Picture = LoadPicture(App.Path + "\SmM_Icos\05.gif")
                      cd = True
                CDid = Val(Left(Right(Lf1.List(Pid), 6), 2))
                        Command = "status cd length track " & CDid
             mciSendString Command, e, Len(e), 0
             Stime = Left(e, 5)
       Info = "[" + Str(Pid + 1) + " -" + Str(Lf1.ListCount) + "]  标题:" + SotPath(Lf1.List(Pid)) + "  艺术家:未知艺术家  唱片集:未知唱片集  流派:未知流派  时间:" + Stime + "  地址:" + Lf1.List(Pid)
        
            If (Playing) Then
                Command = "play cd from " & Val(Left(Right(Lf1.List(Pid), 6), 2))
                SendMCIString Command, True
             Else
                Command = "seek cd to " & Val(Left(Right(Lf1.List(Pid), 6), 2))
                SendMCIString Command, True
                SendMCIString "play cd", True
     End If
           Mper.Filename = "ilxz"
      
                Playing = True
           Tinfo.Left = Pinfo.Width
           
                cd = True

Ig(2).Picture = LoadPicture(App.Path + "\SmM_Icos\05.gif")
If Ly.FileExists(Ly.GetSysPath + "\sndvol32.exe") = True Then
SinfoBo = True
Sinfo
End If

 Me.MousePointer = 0
Exit Sub
End If
Ig(19).Visible = False
Mper.Filename = Lf1.List(Pid)

Me.MousePointer = 0

End Sub

Private Sub Mper_EndOfStream(ByVal result As Long)
On Error Resume Next
Ig(19).Visible = False
Ig(2).Picture = LoadPicture(App.Path + "\SmM_Icos\08.gif")
dfdfvv(0).Enabled = False
dfdfvv(1).Enabled = False
dfdfvv(2).Enabled = False

sdfewe(0).Enabled = False
sdfewe(1).Enabled = False
sdfewe(2).Enabled = False

GoNext

End Sub


Private Sub St1_OnIconEvent(ByVal EventMsg As Long)
If EventMsg = 513 Or EventMsg = 516 Or EventMsg = 519 Then
If Me.WindowState = 1 Then Me.WindowState = LsTy
Me.Show
Me.SetFocus
If EventMsg = 516 Then PopupMenu Me.b002, 40, Screen.Width, Screen.Height
If EventMsg = 519 Then Ig_MouseUp 2, 1, 0, 0, 0

'Ly.MakeTop Me, True
'If a00308.Checked = False Then Ly.MakeTop Me, False


End If

End Sub
Private Sub Timer1_Timer()
On Error Resume Next
If Mper.AnimationAtStart = True Then ReLmp
TX = TX + 1
If SinfoBo = True And TX >= 100 Then
SinfoBo = False
TX = 0
St1.Hidden = True
St1.Hidden = False
End If
If Tinfo.Width > Pinfo.Width Then

If Lor = False Then

  If Ly.FileExists(Ly.GetWinPath + "\sndvol32.exe") = False Then
   Tinfo.Left = Tinfo.Left - 15
  Else
   Tinfo.Left = Tinfo.Left - 30
  End If

   If Tinfo.Left < Pinfo.Width - Tinfo.Width - 500 Then Lor = True



Else
  If Ly.FileExists(Ly.GetWinPath + "\sndvol32.exe") = False Then
   Tinfo.Left = Tinfo.Left + 15
  Else
   Tinfo.Left = Tinfo.Left + 30
  End If
If Tinfo.Left > 500 Then Lor = False




End If



Else
Tinfo.Left = (Pinfo.Width - Tinfo.Width) / 2

End If
If Len(Mper.Filename) > 0 Then
Frame1.Visible = False
Else:
Frame1.Visible = True
End If
If cd = True Then
Update
Frame2.Visible = True
Else
Frame2.Visible = False

End If

If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Change") = True Then LoadSet
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AddFile") = True Then
Lf1.AddItem myReadINI(App.Path + "\SmM_Start.dat", "Start", "AddFile", ""), Lf1.ListCount
AddFile Lf1.List(Lf1.ListCount - 1)

Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AddFile", False
End If
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "PlayFile") = True Then
   Lf1.Clear
                LF2.Clear

   Lf1.AddItem myReadINI(App.Path + "\SmM_Start.dat", "Start", "PlayFile", ""), 0
           AddFile Lf1.List(Lf1.ListCount - 1)

   Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "PlayFile", False
   Lf1.ListIndex = 0
   LF2.ListIndex = 0
   
   LF1_DblClick
  End If
If Mper.DisplaySize = mpFitToSize Then
mjjhjm.Checked = True
a00306.Checked = False
erg.Checked = False
End If


If Len(Mper.Filename) = 0 And cd = False Then Info = "Wellcome to enjoy your digital multimedia with H2O Networks Snowman Media 4se!"
If Txinfo.Caption <> Info Then
Txinfo.Caption = Info
Tinfo.Width = Txinfo.Width
Tinfo.Left = 500
'If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ShowSys") = True Then
'St1.Hidden = True
'St1.Hidden = False
'End If
End If



End Sub

Private Function DirDiver(NewPath As String, dircount As Integer, BackUp As String) As Integer
On Error Resume Next
Static FirstErr As Integer
Dim DirsToPeek As Long, AbandonSearch As Long, INd As Long
Dim OldPath As String, ThePath As String, entry As String
Dim retval As Long
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
    For INd = 0 To File1.ListCount - 1        ' Add conforming files in
        entry = ThePath + File1.List(INd) ' this directory to listbox.
        Lf1.AddItem entry
             AddFile Lf1.List(Lf1.ListCount - 1)
    Next INd
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
    End
  End If
End Function







Private Sub Times_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
Lf1.Clear
             LF2.Clear

For Each ThisFile In Data.Files

Lf1.AddItem ThisFile
        AddFile Lf1.List(Lf1.ListCount - 1)

Next
Pid = 0
Lf1.ListIndex = Pid
LF2.ListIndex = Pid
LF1_DblClick

End Sub


Private Sub Tinfo_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
Lf1.Clear
             LF2.Clear

For Each ThisFile In Data.Files

Lf1.AddItem ThisFile
        AddFile Lf1.List(Lf1.ListCount - 1)

Next
Pid = 0
Lf1.ListIndex = Pid
LF2.ListIndex = Pid
LF1_DblClick


End Sub


Private Sub Txinfo_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ThisFile As Variant
Lf1.Clear
             LF2.Clear

For Each ThisFile In Data.Files

Lf1.AddItem ThisFile
        AddFile Lf1.List(Lf1.ListCount - 1)

Next
Pid = 0
Lf1.ListIndex = Pid
LF2.ListIndex = Pid
LF1_DblClick


End Sub


Private Sub veg_Click()
sdfewer_Click
End Sub

Private Sub werewr_Click()
If Lf1.ListCount > 0 And Lf1.SelCount > 0 Then
If Mper.Filename = Lf1.List(Lf1.ListIndex) Then
Mper.Filename = "ilxz"
End If
Ly.DelFile Lf1.List(Lf1.ListIndex)
If Ly.FileExists(Lf1.List(Lf1.ListIndex)) = False Then
Lf1.RemoveItem (Lf1.ListIndex)
LF2.RemoveItem (Lf1.ListIndex)
  LF2.ListIndex = Pid
If Pid <= Lf1.ListIndex Then Pid = Pid - 1

End If
End If
End Sub

Private Sub xbcty_Click()
Mper.Balance = 0
xbcty.Checked = True
dgv.Checked = True
xbrth.Checked = False
bf6t5.Checked = False
zxvreg.Checked = False
zxvrs.Checked = False
End Sub

Private Sub xbrth_Click()
Mper.Balance = -9640
xbcty.Checked = False
dgv.Checked = False
xbrth.Checked = True
bf6t5.Checked = False
zxvreg.Checked = True
zxvrs.Checked = False

End Sub

Private Sub xfg_Click()
Mper.DisplaySize = mpDoubleSize
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



Private Sub yj6_Click(Index As Integer)
On Error Resume Next
If yj6(0).Checked = True And Index = 0 Then Exit Sub
   Dim rtn As Long
 rtn = GetWindowLong(hwnd, -20)
 rtn = rtn Or &H80000
SetWindowLong hwnd, -20, rtn
For i = 0 To 9
 yj6(i).Checked = False
  asd2(i).Checked = False
Next
yj6(Index).Checked = True
 asd2(Index).Checked = True
If Index = 0 Then SetLayeredWindowAttributes hwnd, 0, 255, &H2
If Index = 1 Then SetLayeredWindowAttributes hwnd, 0, 230, &H2
If Index = 2 Then SetLayeredWindowAttributes hwnd, 0, 204, &H2
If Index = 3 Then SetLayeredWindowAttributes hwnd, 0, 179, &H2
If Index = 4 Then SetLayeredWindowAttributes hwnd, 0, 153, &H2
If Index = 5 Then SetLayeredWindowAttributes hwnd, 0, 128, &H2
If Index = 6 Then SetLayeredWindowAttributes hwnd, 0, 102, &H2
If Index = 7 Then SetLayeredWindowAttributes hwnd, 0, 76, &H2
If Index = 8 Then SetLayeredWindowAttributes hwnd, 0, 51, &H2
If Index = 9 Then SetLayeredWindowAttributes hwnd, 0, 26, &H2
End Sub

Private Sub ytjuytjkuy_Click()
Ig_MouseUp 16, 1, 0, 0, 0
End Sub

Private Sub Update()
On Error Resume Next
Static e As String * 30
mciSendString "status cd media present", e, Len(e), 0
If cd = True Then
     mciSendString "status cd position", e, Len(e), 0
    Track = CInt(Mid$(e, 1, 2))
    Minute = CInt(Mid$(e, 4, 2))
    Second = CInt(Mid$(e, 7, 2))
  
   Times.Caption = Format(Minute, "00") & ":" & Format(Second, "00") + " / " + Left(Stime, 5)
      mciSendString "status cd mode", e, Len(e), 0
    Playing = (Mid$(e, 1, 7) = "playing")
    If Playing = True Then
    Ct.Caption = "正在播放"
    Else
    Ct.Caption = "已暂停"
    
    End If
    Command1.Left = 45 + (Picture1.Width - 280) * (Minute * 60 + Second) / (Int(Left(Stime, 2)) * 60 + Int(Right(Stime, 2)))
    'Else
     'If (CDLoad = True) Then
      '  CDLoad = False
       ' Playing = False
       ' TrackTime.Caption = ""
       ' TimeWindow.Text = ""
   ' End If
End If
If Playing = True Then
If CDid <> Track Then
SendMCIString "stop cd wait", True
Playing = False
  Ct.Caption = "已停止"
  Times.Caption = ""
 Ig(19).Visible = False

cd = False
CDid = 0
GoNext
End If
End If
End Sub

Private Sub ytu5gf_Click()
Ig_MouseUp 11, 1, 0, 0, 0

End Sub

Private Sub zxvreg_Click()
 xbrth_Click
End Sub


Private Sub zxvrs_Click()
bf6t5_Click
End Sub


