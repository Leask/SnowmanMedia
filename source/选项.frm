VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "LYFTOOLS.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sm.M. Setting"
   ClientHeight    =   6390
   ClientLeft      =   2850
   ClientTop       =   1080
   ClientWidth     =   5595
   Icon            =   "选项.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   5595
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3195
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin API控制大全.LyfTools LyfTools1 
      Left            =   4590
      Top             =   6975
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton Command6 
      Caption         =   "取消(&C)"
      Height          =   330
      Left            =   4410
      TabIndex        =   28
      Top             =   5985
      Width           =   1140
   End
   Begin VB.CommandButton Command5 
      Caption         =   "应用(&A)"
      Height          =   330
      Left            =   3150
      TabIndex        =   27
      Top             =   5985
      Width           =   1140
   End
   Begin VB.CommandButton Command4 
      Caption         =   "确定(&Y)"
      Height          =   330
      Left            =   1935
      TabIndex        =   26
      Top             =   5985
      Width           =   1140
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5865
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   10345
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "性能"
      TabPicture(0)   =   "选项.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame13"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Picture1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "外观"
      TabPicture(1)   =   "选项.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label41"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame15"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Skin Window"
      TabPicture(2)   =   "选项.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame14"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame16"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Snowflake"
      TabPicture(3)   =   "选项.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame7"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "同步列表"
      TabPicture(4)   =   "选项.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Image2"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label38"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame12"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame11"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Frame19"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "媒体书签"
      TabPicture(5)   =   "选项.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame3"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame2"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "网络"
      TabPicture(6)   =   "选项.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Image3"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Label39"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Frame9"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Frame10"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "相关组件"
      TabPicture(7)   =   "选项.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Image4"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Label40"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Frame20"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "Command7"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "Command8"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "Command9"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).ControlCount=   6
      Begin VB.CommandButton Command1 
         Caption         =   "重建媒体关联(&R)"
         Height          =   330
         Left            =   1485
         TabIndex        =   112
         Top             =   5355
         Width           =   1815
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   180
         Picture         =   "选项.frx":03EA
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   110
         Top             =   810
         Width           =   510
      End
      Begin VB.Frame Frame7 
         Caption         =   "Snowflake 列表:"
         Height          =   4875
         Left            =   -74820
         TabIndex        =   104
         Top             =   810
         Width           =   1770
         Begin VB.ListBox List3 
            Height          =   3120
            Left            =   135
            TabIndex        =   108
            Top             =   495
            Width           =   1500
         End
         Begin VB.CommandButton Command15 
            Caption         =   "安装(&S)"
            Height          =   330
            Left            =   135
            TabIndex        =   107
            Top             =   4050
            Width           =   1500
         End
         Begin VB.CommandButton Command13 
            Caption         =   "更多(&M)..."
            Height          =   330
            Left            =   135
            TabIndex        =   106
            Top             =   4410
            Width           =   1500
         End
         Begin VB.CommandButton Command11 
            Caption         =   "应用(U)"
            Height          =   330
            Left            =   135
            TabIndex        =   105
            Top             =   3690
            Width           =   1500
         End
         Begin VB.Label Label19 
            Caption         =   "总数:"
            Height          =   195
            Left            =   135
            TabIndex        =   109
            Top             =   270
            Width           =   1365
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "始终位于最前面的窗口:"
         Height          =   915
         Left            =   180
         TabIndex        =   7
         Top             =   4320
         Width           =   5010
         Begin VB.CheckBox Check 
            Caption         =   "Snowflake Window (&S)"
            Height          =   240
            Index           =   9
            Left            =   2070
            TabIndex        =   93
            Top             =   270
            Value           =   1  'Checked
            Width           =   2130
         End
         Begin VB.CheckBox Check 
            Caption         =   "视频窗口(&V)"
            Height          =   240
            Index           =   10
            Left            =   180
            TabIndex        =   9
            Top             =   540
            Value           =   1  'Checked
            Width           =   1770
         End
         Begin VB.CheckBox Check 
            Caption         =   "主窗口(&P)"
            Height          =   240
            Index           =   8
            Left            =   180
            TabIndex        =   8
            Top             =   270
            Width           =   2175
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "基础:"
         Enabled         =   0   'False
         Height          =   645
         Left            =   -74820
         TabIndex        =   85
         Top             =   1440
         Width           =   5010
         Begin VB.CheckBox Check 
            Caption         =   "启用 Skin Window 及 Snowflake 界面(&O)"
            Enabled         =   0   'False
            Height          =   240
            Index           =   13
            Left            =   180
            TabIndex        =   86
            Top             =   270
            Value           =   1  'Checked
            Width           =   4155
         End
      End
      Begin VB.CommandButton Command9 
         Caption         =   "属性(&R)"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -71040
         TabIndex        =   80
         Top             =   3780
         Width           =   1230
      End
      Begin VB.CommandButton Command8 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -72480
         TabIndex        =   79
         Top             =   3780
         Width           =   1230
      End
      Begin VB.CommandButton Command7 
         Caption         =   "添加(&E)"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -73830
         TabIndex        =   78
         Top             =   3780
         Width           =   1230
      End
      Begin VB.Frame Frame20 
         Caption         =   "相关组件列表:"
         Height          =   2220
         Left            =   -74820
         TabIndex        =   76
         Top             =   1440
         Width           =   5010
         Begin VB.ListBox List1 
            Height          =   1740
            ItemData        =   "选项.frx":196C
            Left            =   180
            List            =   "选项.frx":196E
            Style           =   1  'Checkbox
            TabIndex        =   77
            Top             =   270
            Width           =   4650
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "OLE 拖放支持:"
         Height          =   690
         Left            =   -74820
         TabIndex        =   46
         Top             =   2295
         Width           =   5010
         Begin VB.CheckBox Check 
            Caption         =   "启用 OLE 列表拖放支持(&O)"
            Height          =   240
            Index           =   1
            Left            =   180
            TabIndex        =   47
            Top             =   270
            Value           =   1  'Checked
            Width           =   3075
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "所选 Skin:"
         Height          =   4875
         Left            =   -72975
         TabIndex        =   34
         Top             =   810
         Width           =   3165
         Begin VB.Image Image6 
            Height          =   2100
            Left            =   90
            Picture         =   "选项.frx":1970
            Stretch         =   -1  'True
            Top             =   2700
            Width           =   3000
         End
         Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
            Height          =   1995
            Left            =   180
            TabIndex        =   90
            Top             =   2745
            Width           =   2850
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
            ClickToPlay     =   0   'False
            CursorType      =   0
            CurrentPosition =   -1
            CurrentMarker   =   0
            DefaultFrame    =   ""
            DisplayBackColor=   16777215
            DisplayForeColor=   0
            DisplayMode     =   0
            DisplaySize     =   4
            Enabled         =   -1  'True
            EnableContextMenu=   0   'False
            EnablePositionControls=   0   'False
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
            ShowControls    =   0   'False
            ShowAudioControls=   -1  'True
            ShowDisplay     =   0   'False
            ShowGotoBar     =   0   'False
            ShowPositionControls=   0   'False
            ShowStatusBar   =   0   'False
            ShowTracker     =   -1  'True
            TransparentAtStart=   0   'False
            VideoBorderWidth=   0
            VideoBorderColor=   -2147483633
            VideoBorder3D   =   0   'False
            Volume          =   -600
            WindowlessVideo =   0   'False
         End
         Begin VB.Label Label9 
            Caption         =   "备注 : ?"
            Height          =   825
            Left            =   180
            TabIndex        =   40
            Top             =   1665
            Width           =   2850
         End
         Begin VB.Label Label3 
            Caption         =   "时间 : ?"
            Height          =   180
            Left            =   180
            TabIndex        =   36
            Top             =   1440
            Width           =   2900
         End
         Begin VB.Label Label11 
            Caption         =   "站点 : ?"
            Height          =   180
            Left            =   180
            TabIndex        =   41
            Top             =   990
            Width           =   2900
         End
         Begin VB.Label Label12 
            Caption         =   "预览 :"
            Height          =   240
            Left            =   180
            TabIndex        =   42
            Top             =   2520
            Width           =   1680
         End
         Begin VB.Label Label8 
            Caption         =   "版权 : ?"
            Height          =   180
            Left            =   180
            TabIndex        =   39
            Top             =   1215
            Width           =   2900
         End
         Begin VB.Label Label7 
            Caption         =   "公司 : ?"
            Height          =   180
            Left            =   180
            TabIndex        =   38
            Top             =   765
            Width           =   2900
         End
         Begin VB.Label Label5 
            Caption         =   "作者 : ?"
            Height          =   180
            Left            =   180
            TabIndex        =   37
            Top             =   540
            Width           =   2900
         End
         Begin VB.Label Label6 
            Caption         =   "标题 : ?"
            Height          =   180
            Left            =   180
            TabIndex        =   35
            Top             =   270
            Width           =   2900
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Skin 列表:"
         Height          =   4875
         Left            =   -74820
         TabIndex        =   32
         Top             =   810
         Width           =   1770
         Begin VB.CommandButton Command14 
            Caption         =   "应用(U)"
            Height          =   330
            Left            =   135
            TabIndex        =   91
            Top             =   3690
            Width           =   1500
         End
         Begin VB.CommandButton Command12 
            Caption         =   "更多(&M)..."
            Height          =   330
            Left            =   135
            TabIndex        =   88
            Top             =   4410
            Width           =   1500
         End
         Begin VB.CommandButton Command10 
            Caption         =   "安装(&S)"
            Height          =   330
            Left            =   135
            TabIndex        =   82
            Top             =   4050
            Width           =   1500
         End
         Begin VB.ListBox List2 
            Height          =   3120
            Left            =   135
            TabIndex        =   33
            Top             =   495
            Width           =   1500
         End
         Begin VB.Label Label21 
            Caption         =   "总数:"
            Height          =   195
            Left            =   135
            TabIndex        =   89
            Top             =   270
            Width           =   1365
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "重置所有选项(&R)"
         Height          =   330
         Left            =   3375
         TabIndex        =   25
         Top             =   5355
         Width           =   1815
      End
      Begin VB.Frame Frame8 
         Caption         =   "播放机:"
         Height          =   1455
         Left            =   180
         TabIndex        =   23
         Top             =   1440
         Width           =   5010
         Begin VB.CheckBox Check 
            Caption         =   "允许流动网络对你的 Snowman Media 进行唯一识别(&O)"
            Enabled         =   0   'False
            Height          =   240
            Index           =   4
            Left            =   180
            TabIndex        =   24
            Top             =   1080
            Width           =   4740
         End
         Begin VB.CheckBox Check 
            Caption         =   "允许 Snowman Media 自动升级(&L)"
            Enabled         =   0   'False
            Height          =   240
            Index           =   5
            Left            =   180
            TabIndex        =   29
            Top             =   810
            Width           =   4335
         End
         Begin VB.CheckBox Check 
            Caption         =   "允许启动时播放启动音乐(&E)"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   111
            Top             =   270
            Value           =   1  'Checked
            Width           =   4560
         End
         Begin VB.CheckBox Check 
            Caption         =   "允许启动时自动启动断点续播(&F)"
            Height          =   240
            Index           =   14
            Left            =   180
            TabIndex        =   87
            Top             =   540
            Value           =   1  'Checked
            Width           =   4560
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "智能列表整理:"
         Height          =   1365
         Left            =   -74820
         TabIndex        =   20
         Top             =   3150
         Width           =   5010
         Begin VB.OptionButton Option 
            Caption         =   "从不自动进行,直到用户要求(&N)"
            Height          =   285
            Index           =   4
            Left            =   180
            TabIndex        =   45
            Top             =   900
            Width           =   2895
         End
         Begin VB.OptionButton Option 
            Caption         =   "在每次退出 Snowman Media 时自动进行(&E)"
            Height          =   285
            Index           =   3
            Left            =   180
            TabIndex        =   44
            Top             =   585
            Value           =   -1  'True
            Width           =   4470
         End
         Begin VB.OptionButton Option 
            Caption         =   "在每次启动 Snowman Media 时自动进行(&S)"
            Height          =   285
            Index           =   2
            Left            =   180
            TabIndex        =   43
            Top             =   270
            Width           =   4200
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "内容:"
         Height          =   690
         Left            =   -74820
         TabIndex        =   19
         Top             =   1440
         Width           =   5010
         Begin VB.CheckBox Check 
            Caption         =   "退出时清空列表(&D)"
            Height          =   240
            Index           =   2
            Left            =   180
            TabIndex        =   84
            Top             =   270
            Width           =   1905
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "断点续播:"
         Height          =   1680
         Left            =   -74820
         TabIndex        =   18
         Top             =   810
         Width           =   5010
         Begin VB.OptionButton Option 
            Caption         =   "从不自动设置断点,即禁用断点续播(&N)"
            Height          =   240
            Index           =   7
            Left            =   180
            TabIndex        =   50
            Top             =   720
            Width           =   3480
         End
         Begin VB.OptionButton Option 
            Caption         =   "仅在按下停止或退出 Snowman Media 时才设置断点(E)"
            Height          =   240
            Index           =   6
            Left            =   180
            TabIndex        =   49
            Top             =   495
            Value           =   -1  'True
            Width           =   4785
         End
         Begin VB.OptionButton Option 
            Caption         =   "每 5 分钟记录自动断点(&F)"
            Height          =   240
            Index           =   5
            Left            =   180
            TabIndex        =   48
            Top             =   270
            Width           =   2985
         End
         Begin VB.Label Label25 
            Caption         =   "位置     : "
            Height          =   195
            Left            =   225
            TabIndex        =   53
            Top             =   1440
            Width           =   4605
         End
         Begin VB.Label Label24 
            Caption         =   "媒体文件 : "
            Height          =   195
            Left            =   225
            TabIndex        =   52
            Top             =   1215
            Width           =   4605
         End
         Begin VB.Label Label23 
            Caption         =   "当前断点信息:"
            Height          =   240
            Left            =   180
            TabIndex        =   51
            Top             =   990
            Width           =   1995
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "自定义媒体书签信息:"
         Height          =   3120
         Left            =   -74820
         TabIndex        =   17
         Top             =   2565
         Width           =   5010
         Begin VB.Label Label34 
            Caption         =   "位置     : "
            Height          =   240
            Left            =   270
            TabIndex        =   62
            Top             =   2160
            Width           =   4605
         End
         Begin VB.Label Label33 
            Caption         =   "媒体文件 : "
            Height          =   240
            Left            =   270
            TabIndex        =   61
            Top             =   1935
            Width           =   4605
         End
         Begin VB.Label Label37 
            Caption         =   "位置     : "
            Height          =   195
            Left            =   270
            TabIndex        =   65
            Top             =   2880
            Width           =   4605
         End
         Begin VB.Label Label36 
            Caption         =   "媒体文件 : "
            Height          =   240
            Left            =   270
            TabIndex        =   64
            Top             =   2655
            Width           =   4650
         End
         Begin VB.Label Label35 
            Caption         =   "媒体书签 [ C ]:"
            Height          =   240
            Left            =   135
            TabIndex        =   63
            Top             =   1710
            Width           =   4605
         End
         Begin VB.Label Label32 
            Caption         =   "媒体书签 [ D ]:"
            Height          =   240
            Left            =   180
            TabIndex        =   60
            Top             =   2430
            Width           =   4605
         End
         Begin VB.Label Label31 
            Caption         =   "位置     : "
            Height          =   240
            Left            =   270
            TabIndex        =   59
            Top             =   1440
            Width           =   4605
         End
         Begin VB.Label Label30 
            Caption         =   "媒体文件 : "
            Height          =   240
            Left            =   270
            TabIndex        =   58
            Top             =   1215
            Width           =   4605
         End
         Begin VB.Label Label29 
            Caption         =   "媒体书签 [ B ]:"
            Height          =   240
            Left            =   180
            TabIndex        =   57
            Top             =   990
            Width           =   4605
         End
         Begin VB.Label Label28 
            Caption         =   "位置     : "
            Height          =   240
            Left            =   270
            TabIndex        =   56
            Top             =   720
            Width           =   4605
         End
         Begin VB.Label Label27 
            Caption         =   "媒体文件 : "
            Height          =   240
            Left            =   270
            TabIndex        =   55
            Top             =   495
            Width           =   4605
         End
         Begin VB.Label Label26 
            Caption         =   "媒体书签 [ A ]:"
            Height          =   240
            Left            =   180
            TabIndex        =   54
            Top             =   270
            Width           =   4605
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "缓冲:"
         Height          =   1095
         Left            =   -74820
         TabIndex        =   6
         Top             =   4005
         Width           =   5010
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   3105
            TabIndex        =   70
            Text            =   "5"
            Top             =   585
            Width           =   780
         End
         Begin VB.OptionButton Option 
            Caption         =   "指定缓冲区大小为以下秒数(&D)"
            Enabled         =   0   'False
            Height          =   240
            Index           =   13
            Left            =   180
            TabIndex        =   16
            Top             =   585
            Width           =   3525
         End
         Begin VB.OptionButton Option 
            Caption         =   "让 Snowman Media 决定使用缓冲区的大小(&S)"
            Height          =   240
            Index           =   12
            Left            =   180
            TabIndex        =   15
            Top             =   270
            Value           =   -1  'True
            Width           =   3975
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "连接:"
         Height          =   2400
         Left            =   -74820
         TabIndex        =   5
         Top             =   1440
         Width           =   5010
         Begin VB.TextBox Text 
            Height          =   300
            Index           =   1
            Left            =   450
            TabIndex        =   73
            Text            =   "163"
            Top             =   1170
            Width           =   1815
         End
         Begin VB.OptionButton Option 
            Caption         =   "指定以下连接连接网络(&P)"
            Height          =   240
            Index           =   10
            Left            =   180
            TabIndex        =   72
            Top             =   900
            Width           =   2715
         End
         Begin VB.CheckBox Check 
            Caption         =   "启用自动断开功能(&O)"
            Height          =   240
            Index           =   3
            Left            =   180
            TabIndex        =   69
            Top             =   1980
            Width           =   2625
         End
         Begin VB.OptionButton Option 
            Caption         =   "使用网络向导指导连接网络(&W)"
            Height          =   240
            Index           =   11
            Left            =   180
            TabIndex        =   68
            Top             =   1620
            Value           =   -1  'True
            Width           =   3165
         End
         Begin VB.OptionButton Option 
            Caption         =   "拨打默认连接(&M)"
            Height          =   240
            Index           =   9
            Left            =   180
            TabIndex        =   67
            Top             =   585
            Width           =   2805
         End
         Begin VB.OptionButton Option 
            Caption         =   "从不进行拨号连接(&N)"
            Height          =   240
            Index           =   8
            Left            =   180
            TabIndex        =   66
            Top             =   270
            Width           =   2580
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "状态栏:"
         Height          =   915
         Left            =   -74820
         TabIndex        =   4
         Top             =   4770
         Width           =   5010
         Begin VB.OptionButton Option 
            Caption         =   "正常显示(&I)"
            Height          =   258
            Index           =   17
            Left            =   2700
            TabIndex        =   14
            Top             =   540
            Width           =   1365
         End
         Begin VB.OptionButton Option 
            Caption         =   "闪烁效果(&R)"
            Height          =   258
            Index           =   15
            Left            =   2700
            TabIndex        =   13
            Top             =   225
            Width           =   1410
         End
         Begin VB.OptionButton Option 
            Caption         =   "走马灯效果(&G)"
            Height          =   258
            Index           =   16
            Left            =   810
            TabIndex        =   12
            Top             =   540
            Width           =   1500
         End
         Begin VB.OptionButton Option 
            Caption         =   "打字机效果(&L)"
            Height          =   258
            Index           =   14
            Left            =   810
            TabIndex        =   11
            Top             =   225
            Value           =   -1  'True
            Width           =   1500
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "工具栏:"
         Height          =   2400
         Left            =   -74820
         TabIndex        =   3
         Top             =   2205
         Width           =   5010
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00000000&
            Height          =   1230
            Left            =   180
            ScaleHeight     =   1170
            ScaleWidth      =   4590
            TabIndex        =   113
            Top             =   990
            Width           =   4650
         End
         Begin VB.TextBox Text 
            Height          =   300
            Index           =   2
            Left            =   1305
            TabIndex        =   30
            Top             =   585
            Width           =   2445
         End
         Begin VB.CheckBox Check 
            Caption         =   "使用平滑拖动(&S)"
            Height          =   240
            Index           =   11
            Left            =   180
            TabIndex        =   21
            Top             =   270
            Value           =   1  'Checked
            Width           =   1770
         End
         Begin VB.CommandButton Command2 
            Caption         =   "浏览(&B)"
            Height          =   330
            Left            =   3870
            TabIndex        =   10
            Top             =   585
            Width           =   960
         End
         Begin VB.Label Label4 
            Caption         =   "背景更新为:"
            Height          =   240
            Left            =   180
            TabIndex        =   31
            Top             =   630
            Width           =   1500
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "视频:"
         Height          =   1185
         Left            =   180
         TabIndex        =   1
         Top             =   3015
         Width           =   5010
         Begin VB.CheckBox Check 
            Caption         =   "支持鼠标操控(&M)"
            Height          =   240
            Index           =   12
            Left            =   180
            TabIndex        =   92
            Top             =   270
            Width           =   2625
         End
         Begin VB.CheckBox Check 
            Caption         =   "自动禁止屏幕保护"
            Height          =   240
            Index           =   7
            Left            =   180
            TabIndex        =   81
            Top             =   810
            Value           =   1  'Checked
            Width           =   3435
         End
         Begin VB.CheckBox Check 
            Caption         =   "最大化视频窗口时自动进入全屏模式(&W)"
            Height          =   240
            Index           =   6
            Left            =   180
            TabIndex        =   22
            Top             =   540
            Value           =   1  'Checked
            Width           =   4155
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "所选 Snowflake:"
         Height          =   4875
         Left            =   -72975
         TabIndex        =   94
         Top             =   810
         Width           =   3165
         Begin VB.Image Image7 
            Height          =   2100
            Left            =   90
            Picture         =   "选项.frx":1DDD
            Stretch         =   -1  'True
            Top             =   2700
            Width           =   3000
         End
         Begin MediaPlayerCtl.MediaPlayer MediaPlayer2 
            Height          =   1995
            Left            =   180
            TabIndex        =   95
            Top             =   2745
            Width           =   2850
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
            ClickToPlay     =   0   'False
            CursorType      =   0
            CurrentPosition =   -1
            CurrentMarker   =   0
            DefaultFrame    =   ""
            DisplayBackColor=   16777215
            DisplayForeColor=   0
            DisplayMode     =   0
            DisplaySize     =   4
            Enabled         =   -1  'True
            EnableContextMenu=   0   'False
            EnablePositionControls=   0   'False
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
            ShowControls    =   0   'False
            ShowAudioControls=   -1  'True
            ShowDisplay     =   0   'False
            ShowGotoBar     =   0   'False
            ShowPositionControls=   0   'False
            ShowStatusBar   =   0   'False
            ShowTracker     =   -1  'True
            TransparentAtStart=   0   'False
            VideoBorderWidth=   0
            VideoBorderColor=   -2147483633
            VideoBorder3D   =   0   'False
            Volume          =   -600
            WindowlessVideo =   0   'False
         End
         Begin VB.Label Label18 
            Caption         =   "标题 : ?"
            Height          =   180
            Left            =   180
            TabIndex        =   103
            Top             =   270
            Width           =   2900
         End
         Begin VB.Label Label17 
            Caption         =   "作者 : ?"
            Height          =   180
            Left            =   180
            TabIndex        =   102
            Top             =   540
            Width           =   2900
         End
         Begin VB.Label Label16 
            Caption         =   "公司 : ?"
            Height          =   180
            Left            =   180
            TabIndex        =   101
            Top             =   765
            Width           =   2900
         End
         Begin VB.Label Label15 
            Caption         =   "版权 : ?"
            Height          =   180
            Left            =   180
            TabIndex        =   100
            Top             =   1215
            Width           =   2900
         End
         Begin VB.Label Label14 
            Caption         =   "预览 :"
            Height          =   240
            Left            =   180
            TabIndex        =   99
            Top             =   2520
            Width           =   1680
         End
         Begin VB.Label Label13 
            Caption         =   "站点 : ?"
            Height          =   180
            Left            =   180
            TabIndex        =   98
            Top             =   990
            Width           =   2900
         End
         Begin VB.Label Label10 
            Caption         =   "时间 : ?"
            Height          =   180
            Left            =   180
            TabIndex        =   97
            Top             =   1440
            Width           =   2900
         End
         Begin VB.Label Label2 
            Caption         =   "备注 : ?"
            Height          =   825
            Left            =   180
            TabIndex        =   96
            Top             =   1665
            Width           =   2850
         End
      End
      Begin VB.Label Label41 
         Caption         =   "使用本选项卡设置播放机的外观."
         Height          =   240
         Left            =   -74145
         TabIndex        =   83
         Top             =   810
         Width           =   3750
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   -74820
         Picture         =   "选项.frx":224A
         Top             =   810
         Width           =   480
      End
      Begin VB.Label Label40 
         Caption         =   "使用本选项卡查看和配置 Snowman Media 的所有相关组件."
         Height          =   420
         Left            =   -74145
         TabIndex        =   75
         Top             =   810
         Width           =   4380
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   -74820
         Picture         =   "选项.frx":2554
         Top             =   810
         Width           =   480
      End
      Begin VB.Label Label39 
         Caption         =   "使用本选项卡对 Snowman Media 的网络播放功能进行配置."
         Height          =   420
         Left            =   -74145
         TabIndex        =   74
         Top             =   810
         Width           =   4380
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   -74820
         Picture         =   "选项.frx":285E
         Top             =   810
         Width           =   480
      End
      Begin VB.Label Label38 
         Caption         =   "使用本选项卡对 Snowman Media 的同步播放列表进行配置."
         Height          =   420
         Left            =   -74145
         TabIndex        =   71
         Top             =   810
         Width           =   4380
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -74820
         Picture         =   "选项.frx":2B68
         Top             =   810
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "使用本选项卡设置播放机的常规性能选项."
         Height          =   240
         Left            =   855
         TabIndex        =   2
         Top             =   810
         Width           =   3750
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IsOk As Long, pHnd As Long
 Const SYNCHRONIZE = &H100000
 Const INFINITE = &HFFFFFFFF
Private Declare Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long

Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Dim i As Integer
Dim SkiName As String
Dim SkinCount As Integer
Dim txtPath As String
Dim IniPath As String
Dim TmpStr As String
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
Private Function GetFolderValue(wIdx As Integer) As Long
    If wIdx < 2 Then
        GetFolderValue = 0
    ElseIf wIdx < 12 Then
        GetFolderValue = wIdx
    Else
        GetFolderValue = wIdx + 4
    End If
End Function
Sub OPenDir()
  Dim BI As BROWSEINFO
  Dim nFolder As Long
  Dim IDL As ITEMIDLIST
  Dim pIdl As Long
  Dim sPath As String
  Dim SHFI As SHFILEINFO
  Dim m_wCurOptIdx As Integer
  Dim txtDisplayName As String
  Dim noerror, SHGFI_PIDL, SHGFI_ICON, SHGFI_SMALLICON As Integer
   With BI
    .hOwner = Me.hwnd
    nFolder = GetFolderValue(m_wCurOptIdx)
      If SHGetSpecialFolderLocation(ByVal Me.hwnd, ByVal nFolder, IDL) = noerror Then
      .pidlRoot = IDL.mkid.cb
    End If
    .pszDisplayName = String$(MAX_PATH, 0)
    .lpszTitle = "请选配置文件所在的文件夹,选定后 Snowman Media ilxz 3.5 将自动安装."
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
End Sub



Private Sub Command1_Click()
On Error Resume Next
             IsOk = Shell(App.Path + "\SmM_ML.exe", vbNormalFocus)
           pHnd = OpenProcess(SYNCHRONIZE, 0, IsOk)
If pHnd <> 0 Then
    Call WaitForSingleObject(pHnd, INFINITE)
    Call CloseHandle(pHnd)

End If

End Sub

'Private Sub Command1_Click()
'Form102.MediaPlayer1.ShowDialog mpShowDialogOptions
'End Sub
Private Sub Command10_Click()
      IniPath = App.Path + "\smm_ssct.ini"
 Call OPenDir
     If myReadINI(txtPath + "\skin_info.skin", "used", "used", "") = 0 Then
               TmpStr = myReadINI(txtPath + "\skin_info.skin", "info", "title", "")
            List2.AddItem TmpStr
       LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_" + Str(myReadINI(IniPath, "skin", "count", "") + 1) + "_Name", myReadINI(txtPath + "\skin_info.skin", "info", "title", "")
       LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_" + Str(myReadINI(IniPath, "skin", "count", "") + 1) + "_Path", txtPath
         Dim i As Integer
         i = myWriteINI(App.Path + "\SmM_SSCT.ini", "skin", "count", List2.ListCount)
         Form_Load
       GetInfo (List2.ListCount)
    Else:
   If Len(txtPath) > 0 Then
   MsgBox ("无法在指定的路径下找到完整的 Skin 配置文件,该 Skin 可能已经损坏.请确定路径正确和 Skin 完整后重试.")
   End If
   End If
            
            
             
End Sub
Private Sub Command15_Click()
      IniPath = App.Path + "\smm_ssct.ini"
 Call OPenDir
     If myReadINI(txtPath + "\sflake_info.sfl", "used", "used", "") = 0 Then
               TmpStr = myReadINI(txtPath + "\sflake_info.sfl", "info", "title", "")
            List3.AddItem TmpStr
       LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_" + Str(myReadINI(IniPath, "sflake", "count", "") + 1) + "_Name", myReadINI(txtPath + "\sflake_info.sfl", "info", "title", "")
       LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_" + Str(myReadINI(IniPath, "sflake", "count", "") + 1) + "_Path", txtPath
         
         Dim i As Integer
         i = myWriteINI(App.Path + "\SmM_SSCT.ini", "sflake", "count", List3.ListCount)
         Form_Load
       GetInfo (List3.ListCount)
    Else:
   If Len(txtPath) > 0 Then MsgBox ("无法在指定的路径下找到完整的 Snowflake 配置文件,该 Snowflake 可能已经损坏.请确定路径正确和 Skin 完整后重试.")
   End If
            
  
            
            
             
End Sub

Private Sub Command12_Click()
LyfTools1.HttpTo ("http://www.h2ont.com/SnowmanMedia/SkAndAf.html")
End Sub

Private Sub Command13_Click()
LyfTools1.HttpTo ("http://www.h2ont.com/SnowmanMedia/SkAndAf.html")
End Sub

Private Sub Command14_Click()
If Len(List2.List(List2.ListIndex)) > 0 Then
GetInfo (List2.ListIndex + 1)
LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path", LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_" + Str(List2.ListIndex + 1) + "_Path")
'Form102.SkinForm1.SkinPath = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path")

Else: MsgBox ("没有选定的 Skin 无法应用,请选择好要应用的外观后重试.")
End If
End Sub
Private Sub Command11_Click()
If Len(List3.List(List3.ListIndex)) > 0 Then
GetInfoB (List3.ListIndex + 1)
LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path", LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_" + Str(List3.ListIndex + 1) + "_Path")
'Form102.SkinForm1.SkinPath = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path")

Else: MsgBox ("没有选定的 Snowflake 无法应用,请选择好要应用的外观后重试.")
End If
End Sub

Private Sub Command2_Click()
    CommonDialog1.Filter = "图片文件:Bmp、Bpg、Did、Wmf、Ico、Gif、Rle、Cur、Emf、Png" & _
    "|*.bmp;*.jpg;*.did;*.wmf;*.ico;*.gif;*.rle;*.cur;*.emf;*.png|所有文件:*.*|*.*"
        CommonDialog1.Filename = ""
    CommonDialog1.ShowOpen
    If Len(CommonDialog1.Filename) > 0 Then
      Text(2).Text = CommonDialog1.Filename
      Picture3.Picture = LoadPicture(Text(2).Text)
   End If
End Sub

Private Sub Command3_Click()
Check(14).Value = 1
Check(0).Value = 1
Check(1).Value = 1
Check(2).Value = 0
Check(4).Value = 0
Check(5).Value = 0
Check(12).Value = 0
Check(6).Value = 1
Check(7).Value = 1
Check(8).Value = 0
Check(9).Value = 1
Check(10).Value = 1
Check(11).Value = 1
Check(3).Value = 0
Check(13).Value = 1
Me.Option(14).Value = 1
Me.Option(15).Value = 0
Me.Option(16).Value = 0
Me.Option(17).Value = 0

Me.Option(2).Value = 1

Me.Option(4).Value = 0
Me.Option(3).Value = 1
Me.Option(5).Value = 0
Me.Option(7).Value = 0
Me.Option(6).Value = 1
Me.Option(8).Value = 0
Me.Option(9).Value = 0
Me.Option(10).Value = 0
Me.Option(11).Value = 1
Me.Option(12).Value = 1
Me.Option(13).Value = 0
Text(2).Text = App.Path + "\BPics\BCG_104.gif"
Text(1).Text = "163"
Call Regedit
Form_Load
End Sub
Private Sub Command4_Click()
Call Regedit
MsgBox ("所作的修改将在下次启动 Sm.M. 时生效.")
End
End Sub
Private Sub Command5_Click()
MsgBox ("所作的修改将在下次启动 Sm.M. 时生效.")
Call Regedit
End Sub
Private Sub Command6_Click()
End
End Sub
Sub Regedit()
For i = 2 To 17
LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Op_" + Str(i), Me.Option(i).Value
Next
For i = 0 To 14
LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(i), Me.Check(i).Value
Next
For i = 1 To 2
LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Te_" + Str(i), Me.Text(i).Text
Next
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim test As String
List2.Clear
List3.Clear
Open App.Path + "\SmM_OD.dat" For Input As #1
    While Not EOF(1)
    Line Input #1, test
    If test <> "" Then List1.AddItem RTrim(test)
    Wend
    Close #1
Text(2).Text = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Te_ 2")
IniPath = App.Path + "\SmM_SSCT.INI"
Label21.Caption = "总数:" + Str(myReadINI(IniPath, "skin", "count", ""))
Label19.Caption = "总数:" + Str(myReadINI(IniPath, "sflake", "count", ""))
For i = 0 To List1.ListCount - 1
List1.Selected(i) = True
Next
List1.ListIndex = 0
For i = 2 To 17
Me.Option(i).Value = LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Op_" + Str(i))
Next
For i = 0 To 14
Me.Check(i).Value = LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(i))
Next
For i = 1 To 2
 Me.Text(i).Text = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Te_" + Str(i))

Next
  Picture3.Picture = LoadPicture(Text(2).Text)
  For i = 1 To myReadINI(IniPath, "skin", "count", "")
   If LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_" + Str(i) + "_Name") <> "Error" Then
   List2.AddItem LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_" + Str(i) + "_Name")
   End If
Next
  For i = 1 To myReadINI(IniPath, "sflake", "count", "")
   If LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_" + Str(i) + "_Name") <> "Error" Then
   List3.AddItem LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_" + Str(i) + "_Name")
   End If
Next
Label24.Caption = "媒体文件 : " + LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name")
Label25.Caption = "位置     : " + LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute")

Label27.Caption = "媒体文件 : " + LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_A")
Label28.Caption = "位置     : " + LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_A")
Label30.Caption = "媒体文件 : " + LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_B")
Label31.Caption = "位置     : " + LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_B")
Label33.Caption = "媒体文件 : " + LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_C")
Label34.Caption = "位置     : " + LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_C")
Label36.Caption = "媒体文件 : " + LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_D")
Label37.Caption = "位置     : " + LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_D")

If Len(LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name")) > 25 Then
Label24.Caption = "媒体文件 : ..." + Right(LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name"), 22)
End If

If Len(LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_A")) > 25 Then
Label27.Caption = "媒体文件 : ..." + Right(LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_A"), 22)
End If
If Len(LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_B")) > 25 Then
Label30.Caption = "媒体文件 : ..." + Right(LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_B"), 22)
End If
If Len(LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_C")) > 25 Then
Label33.Caption = "媒体文件 : ..." + Right(LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_C"), 22)
End If
If Len(LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_D")) > 25 Then
Label36.Caption = "媒体文件 : ..." + Right(LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_D"), 22)
End If


If Me.Option(10).Value = True Then Text(1).Enabled = True
If Me.Option(10).Value = False Then Text(1).Enabled = False
IniPath = App.Path + "\SmM_SSCT.INI"

For i = 1 To myReadINI(IniPath, "skin", "count", "")
If LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path") = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_" + Str(i) + "_Path") Then
List2.ListIndex = i - 1
GetInfo (i)
Exit For
End If
Next
IniPath = App.Path + "\SmM_SSCT.INI"

For i = 1 To myReadINI(IniPath, "sflake", "count", "")
If LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path") = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_" + Str(i) + "_Path") Then
List3.ListIndex = i - 1
GetInfoB (i)
Exit For
End If
Next
End Sub

Private Sub List1_Click()
For i = 0 To List1.ListCount - 1
List1.Selected(i) = 1
Next
End Sub

Private Sub List2_Click()
GetInfo (List2.ListIndex + 1)
End Sub
Private Sub List3_Click()
GetInfoB (List3.ListIndex + 1)
End Sub

Private Sub List2_DblClick()
If Len(List2.List(List2.ListIndex)) > 0 Then
GetInfo (List2.ListIndex + 1)
LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path", LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_" + Str(List2.ListIndex + 1) + "_Path")
'Form102.SkinForm1.SkinPath = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path")
'Me.Show
End If
End Sub
Private Sub List3_DblClick()
If Len(List3.List(List3.ListIndex)) > 0 Then
GetInfoB (List3.ListIndex + 1)
LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path", LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_" + Str(List3.ListIndex + 1) + "_Path")
'Form102.SkinForm1.SkinPath = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path")
'Me.Show
End If
End Sub
Private Sub Option_Click(Index As Integer)
    Select Case Index
             Case 10
               Text(1).Enabled = True
             Case 8
               Text(1).Enabled = False
             Case 9
               Text(1).Enabled = False
             Case 11
               Text(1).Enabled = False
    End Select
End Sub
Private Function GetInfo(No As Integer)
IniPath = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_" + Str(No) + "_Path")
Label6.Caption = "标题 : " + myReadINI(IniPath + "\skin_info.skin", "info", "title", "")
Label5.Caption = "作者 : " + myReadINI(IniPath + "\skin_info.skin", "info", "artist", "")
Label7.Caption = "公司 : " + myReadINI(IniPath + "\skin_info.skin", "info", "con", "")
Label8.Caption = "版权 : " + myReadINI(IniPath + "\skin_info.skin", "info", "copy", "")
Label3.Caption = "时间 : " + myReadINI(IniPath + "\skin_info.skin", "info", "time", "")
Label11.Caption = "站点 : " + myReadINI(IniPath + "\skin_info.skin", "info", "web", "")
Label9.Caption = "备注 : " + myReadINI(IniPath + "\skin_info.skin", "info", "info", "")
MediaPlayer1.Filename = IniPath + "\skin_info.bmp"
End Function
Private Function GetInfoB(No As Integer)
IniPath = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_" + Str(No) + "_Path")
Label18.Caption = "标题 : " + myReadINI(IniPath + "\sflake_info.sfl", "info", "title", "")
Label17.Caption = "作者 : " + myReadINI(IniPath + "\sflake_info.sfl", "info", "artist", "")
Label16.Caption = "公司 : " + myReadINI(IniPath + "\sflake_info.sfl", "info", "con", "")
Label15.Caption = "版权 : " + myReadINI(IniPath + "\sflake_info.sfl", "info", "copy", "")
Label10.Caption = "时间 : " + myReadINI(IniPath + "\sflake_info.sfl", "info", "time", "")
Label13.Caption = "站点 : " + myReadINI(IniPath + "\sflake_info.sfl", "info", "web", "")
Label2.Caption = "备注 : " + myReadINI(IniPath + "\sflake_info.sfl", "info", "info", "")
MediaPlayer2.Filename = IniPath + "\sflake_info.bmp"
End Function


