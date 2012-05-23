VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Object = "{244E6785-6684-11D2-943F-A976CFB4FC0C}#1.0#0"; "CTLSTBAR.OCX"
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "LYFTOOLS.OCX"
Object = "{C40E7B9F-6CF0-11D2-AA70-444553540000}#1.0#0"; "COOLINEPRJ.OCX"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{CFCDAA00-8BE4-11CF-B84B-0020AFBBCCFA}#1.0#0"; "RMOC3260.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form102 
   Caption         =   "Snowman Media  3.0"
   ClientHeight    =   5940
   ClientLeft      =   4200
   ClientTop       =   2325
   ClientWidth     =   6120
   ForeColor       =   &H00000000&
   Icon            =   "Cd Player12.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5940
   ScaleWidth      =   6120
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1665
      TabIndex        =   110
      Top             =   495
      Width           =   1005
   End
   Begin VB.ComboBox cbomonth 
      Height          =   300
      Left            =   2115
      Style           =   2  'Dropdown List
      TabIndex        =   105
      Top             =   7380
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ComboBox cboyear 
      Height          =   300
      Left            =   3375
      Style           =   2  'Dropdown List
      TabIndex        =   104
      Top             =   7425
      Visible         =   0   'False
      Width           =   1215
   End
   Begin CooLinePrj.CooLine CooLine1 
      Height          =   285
      Left            =   45
      TabIndex        =   74
      Top             =   45
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   503
      InsChr          =   95
      Display         =   "Leask Snowman Media  3.0 Ready  - H2ont Running Splerdors!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   65535
      ForeColor       =   16711680
   End
   Begin VB.Frame Frame11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Caption         =   "Frame10"
      ForeColor       =   &H80000008&
      Height          =   5460
      Left            =   6525
      TabIndex        =   72
      Top             =   405
      Width           =   4515
      Begin VB.PictureBox Picture26 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2700
         ScaleHeight     =   255
         ScaleWidth      =   1785
         TabIndex        =   99
         Top             =   4635
         Width           =   1815
         Begin VB.Label Label43 
            BackColor       =   &H0000FFFF&
            Caption         =   "从库中导出列表(&O)"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   45
            TabIndex        =   100
            Top             =   45
            Width           =   2625
         End
      End
      Begin VB.PictureBox Picture25 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2700
         ScaleHeight     =   255
         ScaleWidth      =   1785
         TabIndex        =   97
         Top             =   4365
         Width           =   1815
         Begin VB.Label Label42 
            BackColor       =   &H0000FFFF&
            Caption         =   "导入列表到库中(&I)"
            ForeColor       =   &H00FF0000&
            Height          =   465
            Left            =   45
            TabIndex        =   98
            Top             =   45
            Width           =   2985
         End
      End
      Begin VB.PictureBox Picture24 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2700
         ScaleHeight     =   255
         ScaleWidth      =   1785
         TabIndex        =   95
         Top             =   4095
         Width           =   1815
         Begin VB.Label Label41 
            BackColor       =   &H0000FFFF&
            Caption         =   "从库中删除媒体(&D)"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   45
            TabIndex        =   96
            Top             =   45
            Width           =   2625
         End
      End
      Begin VB.PictureBox Picture22 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2700
         ScaleHeight     =   255
         ScaleWidth      =   1785
         TabIndex        =   93
         Top             =   3825
         Width           =   1815
         Begin VB.Label Label40 
            BackColor       =   &H0000FFFF&
            Caption         =   "添加媒体到库中(&A)"
            ForeColor       =   &H00FF0000&
            Height          =   465
            Left            =   45
            TabIndex        =   94
            Top             =   45
            Width           =   2985
         End
      End
      Begin VB.PictureBox Picture23 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   3795
         Left            =   2700
         ScaleHeight     =   3765
         ScaleWidth      =   1785
         TabIndex        =   81
         Top             =   0
         Width           =   1815
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            Caption         =   "Xxxxxxx"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   315
            TabIndex        =   106
            Top             =   1485
            Width           =   630
         End
         Begin VB.Label Label39 
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   1410
            Left            =   315
            TabIndex        =   92
            Top             =   2295
            Width           =   1320
         End
         Begin VB.Label Label38 
            BackColor       =   &H0080FFFF&
            Caption         =   "更新备注:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   90
            TabIndex        =   91
            Top             =   2070
            Width           =   1680
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            Caption         =   "xxxx"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   315
            TabIndex        =   90
            Top             =   675
            Width           =   360
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            Caption         =   "xx.xx.xxxx"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   315
            TabIndex        =   89
            Top             =   1305
            Width           =   900
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            Caption         =   "xx:xx"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   315
            TabIndex        =   88
            Top             =   1665
            Width           =   450
         End
         Begin VB.Label Label35 
            BackColor       =   &H0080FFFF&
            Caption         =   "最近更新时间:"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   90
            TabIndex        =   87
            Top             =   1080
            Width           =   1500
         End
         Begin VB.Label Label34 
            BackColor       =   &H0080FFFF&
            Caption         =   "[ 媒体库信息 ]"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   90
            TabIndex        =   86
            Top             =   90
            Width           =   1770
         End
         Begin VB.Label Label33 
            BackColor       =   &H0080FFFF&
            Caption         =   "媒体曲目总数:"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   90
            TabIndex        =   85
            Top             =   450
            Width           =   2040
         End
      End
      Begin VB.PictureBox Picture21 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2700
         ScaleHeight     =   255
         ScaleWidth      =   1785
         TabIndex        =   80
         Top             =   4905
         Width           =   1815
         Begin VB.Label Label31 
            BackColor       =   &H0000FFFF&
            Caption         =   "播放所有媒体(&P)"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   45
            TabIndex        =   84
            Top             =   45
            Width           =   1680
         End
      End
      Begin VB.PictureBox Picture20 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1350
         ScaleHeight     =   255
         ScaleWidth      =   1290
         TabIndex        =   79
         Top             =   4905
         Width           =   1320
         Begin VB.Label Label30 
            BackColor       =   &H0000FFFF&
            Caption         =   "停止更新(&S)"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   45
            TabIndex        =   83
            Top             =   45
            Width           =   2625
         End
      End
      Begin VB.PictureBox Picture19 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   78
         Top             =   4905
         Width           =   1365
         Begin VB.Label Label29 
            BackColor       =   &H0000FFFF&
            Caption         =   "更新媒体库(&R)"
            ForeColor       =   &H00FF0000&
            Height          =   465
            Left            =   45
            TabIndex        =   82
            Top             =   45
            Width           =   2985
         End
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   4890
         ItemData        =   "Cd Player12.frx":1582
         Left            =   0
         List            =   "Cd Player12.frx":1584
         OLEDropMode     =   1  'Manual
         Sorted          =   -1  'True
         TabIndex        =   77
         Top             =   0
         Width           =   2670
      End
      Begin VB.PictureBox Picture18 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   4485
         TabIndex        =   75
         Top             =   5175
         Width           =   4515
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0080FFFF&
            Caption         =   "Snowman Media Warehouse  1.0  Ready"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   45
            TabIndex        =   76
            Top             =   45
            Width           =   3150
         End
      End
   End
   Begin ACTIVESKINLibCtl.SkinForm SkinForm1 
      Height          =   480
      Left            =   1575
      OleObjectBlob   =   "Cd Player12.frx":1586
      TabIndex        =   66
      Top             =   7380
      Visible         =   0   'False
      Width           =   480
   End
   Begin API控制大全.LyfTools LyfTools1 
      Left            =   360
      Top             =   7290
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin CTLISTBARLibCtl.ctListBar ctListBar1 
      Height          =   5460
      Left            =   45
      TabIndex        =   21
      ToolTipText     =   "功能菜单  - Snowman Media  3.0"
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
      BackImage       =   "Cd Player12.frx":15CF
      ListForeColor   =   16711680
      BarForeColor    =   16711680
      WordWrap        =   -1  'True
      Caption         =   "正在播放"
      PicArray0       =   "Cd Player12.frx":4694
      PicArray1       =   "Cd Player12.frx":49AE
      PicArray2       =   "Cd Player12.frx":4CC8
      PicArray3       =   "Cd Player12.frx":511A
      PicArray4       =   "Cd Player12.frx":5434
      PicArray5       =   "Cd Player12.frx":574E
      PicArray6       =   "Cd Player12.frx":5A68
      PicArray7       =   "Cd Player12.frx":5D82
      PicArray8       =   "Cd Player12.frx":609C
      PicArray9       =   "Cd Player12.frx":6976
      PicArray10      =   "Cd Player12.frx":6C90
      PicArray11      =   "Cd Player12.frx":6FAA
      PicArray12      =   "Cd Player12.frx":72C4
      PicArray13      =   "Cd Player12.frx":75DE
      PicArray14      =   "Cd Player12.frx":78F8
      PicArray15      =   "Cd Player12.frx":7C12
      PicArray16      =   "Cd Player12.frx":7F2C
      PicArray17      =   "Cd Player12.frx":8246
      PicArray18      =   "Cd Player12.frx":8560
      PicArray19      =   "Cd Player12.frx":89B2
      PicArray20      =   "Cd Player12.frx":8CCC
      PicArray21      =   "Cd Player12.frx":8FE6
      PicArray22      =   "Cd Player12.frx":9300
      PicArray23      =   "Cd Player12.frx":961A
      PicArray24      =   "Cd Player12.frx":9A6C
      PicArray25      =   "Cd Player12.frx":9D86
      PicArray26      =   "Cd Player12.frx":A0A0
      PicArray27      =   "Cd Player12.frx":B632
      PicArray28      =   "Cd Player12.frx":B94C
      PicArray29      =   "Cd Player12.frx":BC66
      PicArray30      =   "Cd Player12.frx":BF80
      PicArray31      =   "Cd Player12.frx":C29A
      PicArray32      =   "Cd Player12.frx":C5B4
      PicArray33      =   "Cd Player12.frx":C8CE
      PicArray34      =   "Cd Player12.frx":CBE8
      PicArray35      =   "Cd Player12.frx":CF02
      PicArray36      =   "Cd Player12.frx":D21C
      PicArray37      =   "Cd Player12.frx":D536
      PicArray38      =   "Cd Player12.frx":DE10
      PicArray39      =   "Cd Player12.frx":E12A
      PicArray40      =   "Cd Player12.frx":E444
      PicArray41      =   "Cd Player12.frx":F6C6
      PicArray42      =   "Cd Player12.frx":F9E0
      PicArray43      =   "Cd Player12.frx":FE32
      PicArray44      =   "Cd Player12.frx":1014C
      PicArray45      =   "Cd Player12.frx":10466
      PicArray46      =   "Cd Player12.frx":10780
      PicArray47      =   "Cd Player12.frx":10A9A
      PicArray48      =   "Cd Player12.frx":10DB4
      PicArray49      =   "Cd Player12.frx":110CE
      PicArray50      =   "Cd Player12.frx":113E8
      PicArray51      =   "Cd Player12.frx":11404
      PicArray52      =   "Cd Player12.frx":11420
      PicArray53      =   "Cd Player12.frx":1143C
      PicArray54      =   "Cd Player12.frx":11458
      PicArray55      =   "Cd Player12.frx":11474
      PicArray56      =   "Cd Player12.frx":11490
      PicArray57      =   "Cd Player12.frx":114AC
      PicArray58      =   "Cd Player12.frx":114C8
      PicArray59      =   "Cd Player12.frx":114E4
      PicArray60      =   "Cd Player12.frx":11500
      PicArray61      =   "Cd Player12.frx":1151C
      PicArray62      =   "Cd Player12.frx":11538
      PicArray63      =   "Cd Player12.frx":11554
      PicArray64      =   "Cd Player12.frx":11570
      PicArray65      =   "Cd Player12.frx":1158C
      PicArray66      =   "Cd Player12.frx":115A8
      PicArray67      =   "Cd Player12.frx":115C4
      PicArray68      =   "Cd Player12.frx":115E0
      PicArray69      =   "Cd Player12.frx":115FC
      PicArray70      =   "Cd Player12.frx":11618
      PicArray71      =   "Cd Player12.frx":11634
      PicArray72      =   "Cd Player12.frx":11650
      PicArray73      =   "Cd Player12.frx":1166C
      PicArray74      =   "Cd Player12.frx":11688
      PicArray75      =   "Cd Player12.frx":116A4
      PicArray76      =   "Cd Player12.frx":116C0
      PicArray77      =   "Cd Player12.frx":116DC
      PicArray78      =   "Cd Player12.frx":116F8
      PicArray79      =   "Cd Player12.frx":11714
      PicArray80      =   "Cd Player12.frx":11730
      PicArray81      =   "Cd Player12.frx":1174C
      PicArray82      =   "Cd Player12.frx":11768
      PicArray83      =   "Cd Player12.frx":11784
      PicArray84      =   "Cd Player12.frx":117A0
      PicArray85      =   "Cd Player12.frx":117BC
      PicArray86      =   "Cd Player12.frx":117D8
      PicArray87      =   "Cd Player12.frx":117F4
      PicArray88      =   "Cd Player12.frx":11810
      PicArray89      =   "Cd Player12.frx":1182C
      PicArray90      =   "Cd Player12.frx":11848
      PicArray91      =   "Cd Player12.frx":11864
      PicArray92      =   "Cd Player12.frx":11880
      PicArray93      =   "Cd Player12.frx":1189C
      PicArray94      =   "Cd Player12.frx":118B8
      PicArray95      =   "Cd Player12.frx":118D4
      PicArray96      =   "Cd Player12.frx":118F0
      PicArray97      =   "Cd Player12.frx":1190C
      PicArray98      =   "Cd Player12.frx":11928
      PicArray99      =   "Cd Player12.frx":11944
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   915
      Left            =   1530
      TabIndex        =   13
      Top             =   4950
      Width           =   4515
      Begin VB.VScrollBar Volume 
         Height          =   915
         Left            =   4365
         MouseIcon       =   "Cd Player12.frx":11960
         TabIndex        =   15
         Top             =   0
         Width           =   150
      End
      Begin VB.ListBox ListFile 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   930
         ItemData        =   "Cd Player12.frx":11AB2
         Left            =   0
         List            =   "Cd Player12.frx":11AB4
         OLEDropMode     =   1  'Manual
         TabIndex        =   14
         ToolTipText     =   "当前播放列表  - Snowman Media  3.0"
         Top             =   0
         Width           =   4380
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4470
      Left            =   1530
      TabIndex        =   4
      Top             =   405
      Width           =   4515
      Begin RealAudioObjectsCtl.RealAudio RA 
         Height          =   3390
         Left            =   10000
         TabIndex        =   109
         Top             =   0
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
         Left            =   2880
         MouseIcon       =   "Cd Player12.frx":11AB6
         TabIndex        =   9
         Text            =   "S.M."
         ToolTipText     =   "CD曲目选择  - Snowman Media  3.0"
         Top             =   3735
         Width           =   780
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2880
         TabIndex        =   18
         Top             =   3690
         Width           =   870
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "随机"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   0
            MaskColor       =   &H00FF0000&
            TabIndex        =   20
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
            TabIndex        =   19
            Top             =   225
            Width           =   690
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   45
         TabIndex        =   5
         Top             =   4185
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
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "[00]00:00"
            ToolTipText     =   "播放时间  - Snowman Media  2.0"
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
            TabIndex        =   8
            ToolTipText     =   "总时间  - Snowman Media  2.0"
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
            TabIndex        =   7
            ToolTipText     =   "本首时间  - Snowman Media  2.0"
            Top             =   0
            Width           =   90
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   3690
         Width           =   5190
         Begin VB.Image Image16 
            Height          =   375
            Left            =   3780
            Picture         =   "Cd Player12.frx":11C08
            Stretch         =   -1  'True
            ToolTipText     =   "媒体播放  － Snowman Media  3.0"
            Top             =   0
            Width           =   345
         End
         Begin VB.Image Image17 
            Height          =   420
            Left            =   4185
            Picture         =   "Cd Player12.frx":12007
            Stretch         =   -1  'True
            ToolTipText     =   "CD播放  － Snowman Media  3.0"
            Top             =   0
            Width           =   435
         End
         Begin VB.Image Image5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   675
            MouseIcon       =   "Cd Player12.frx":12406
            Picture         =   "Cd Player12.frx":12558
            Stretch         =   -1  'True
            ToolTipText     =   "上一首曲目  - Snowman Media  3.0"
            Top             =   0
            Width           =   300
         End
         Begin VB.Image Image7 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1395
            MouseIcon       =   "Cd Player12.frx":12957
            Picture         =   "Cd Player12.frx":12AA9
            Stretch         =   -1  'True
            ToolTipText     =   "快进  - Snowman Media  3.0"
            Top             =   0
            Width           =   300
         End
         Begin VB.Image Image11 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2475
            MouseIcon       =   "Cd Player12.frx":12EA8
            Picture         =   "Cd Player12.frx":12FFA
            Stretch         =   -1  'True
            ToolTipText     =   "弹出光驱  - Snowman Media  3.0"
            Top             =   90
            Width           =   300
         End
         Begin VB.Image Image4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1755
            MouseIcon       =   "Cd Player12.frx":133F9
            Picture         =   "Cd Player12.frx":1354B
            Stretch         =   -1  'True
            ToolTipText     =   "下一首曲目  - Snowman Media  2.0"
            Top             =   90
            Width           =   300
         End
         Begin VB.Image Image3 
            Appearance      =   0  'Flat
            Height          =   240
            Left            =   360
            MouseIcon       =   "Cd Player12.frx":1394A
            Picture         =   "Cd Player12.frx":13A9C
            Stretch         =   -1  'True
            ToolTipText     =   "暂停播放  - Snowman Media  3.0"
            Top             =   135
            Width           =   255
         End
         Begin VB.Image Image18 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   0
            MouseIcon       =   "Cd Player12.frx":13E9B
            Picture         =   "Cd Player12.frx":13FED
            Stretch         =   -1  'True
            ToolTipText     =   "开始播放  - Snowman Media  3.0"
            Top             =   0
            Width           =   345
         End
         Begin VB.Image Image6 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1035
            MouseIcon       =   "Cd Player12.frx":143EC
            Picture         =   "Cd Player12.frx":1453E
            Stretch         =   -1  'True
            ToolTipText     =   "快退  - Snowman Media  3.0"
            Top             =   90
            Width           =   300
         End
         Begin VB.Image Image2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2115
            MouseIcon       =   "Cd Player12.frx":1493D
            Picture         =   "Cd Player12.frx":14A8F
            Stretch         =   -1  'True
            ToolTipText     =   "停止播放  - Snowman Media  3.0"
            Top             =   0
            Width           =   300
         End
         Begin VB.Image Image15 
            Height          =   780
            Left            =   0
            Picture         =   "Cd Player12.frx":14E8E
            Top             =   0
            Width           =   4515
         End
      End
      Begin MSComctlLib.Slider SLD1 
         Height          =   330
         Left            =   10000
         TabIndex        =   111
         Top             =   3420
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   582
         _Version        =   393216
         Max             =   1000
         SelectRange     =   -1  'True
         TickStyle       =   3
      End
      Begin VB.Image Image10 
         Height          =   3375
         Left            =   0
         Picture         =   "Cd Player12.frx":15F19
         Top             =   0
         Width           =   4500
      End
      Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
         DragIcon        =   "Cd Player12.frx":18064
         Height          =   4470
         Left            =   0
         TabIndex        =   11
         Top             =   0
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
         PlayCount       =   0
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
         SendMouseMoveEvents=   0   'False
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
   Begin VB.CommandButton Eject 
      Appearance      =   0  'Flat
      Caption         =   "Open"
      Enabled         =   0   'False
      Height          =   435
      Left            =   945
      TabIndex        =   0
      ToolTipText     =   "Eject CD"
      Top             =   7335
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1215
      Top             =   6795
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Snowman Media  3.0"
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5460
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   1365
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   4470
      Left            =   1575
      TabIndex        =   12
      Top             =   450
      Width           =   4515
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   915
      Left            =   1575
      TabIndex        =   16
      Top             =   4995
      Width           =   4515
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   5460
      Left            =   11205
      TabIndex        =   22
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
         TabIndex        =   24
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
            TabIndex        =   31
            Top             =   0
            Width           =   1095
            Begin VB.Label Label15 
               BackColor       =   &H0000FFFF&
               Caption         =   "清除(&L)"
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   45
               TabIndex        =   32
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
            TabIndex        =   29
            Top             =   4320
            Width           =   1095
            Begin VB.Label Label17 
               BackColor       =   &H0000FFFF&
               Caption         =   "取消(&C)"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   45
               TabIndex        =   30
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
            TabIndex        =   27
            Top             =   270
            Width           =   1365
            Begin VB.Label Label19 
               BackColor       =   &H0000FFFF&
               Caption         =   "当前曲目(&T)"
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   45
               TabIndex        =   28
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
            TabIndex        =   25
            Top             =   540
            Width           =   1365
            Begin VB.Label Label20 
               BackColor       =   &H0000FFFF&
               Caption         =   "播放列表(&L)"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   45
               TabIndex        =   26
               Top             =   45
               Width           =   1635
            End
         End
         Begin VB.Label Label6 
            BackColor       =   &H0000FFFF&
            Caption         =   "预览(&P)"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   90
            TabIndex        =   33
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
         TabIndex        =   50
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
            TabIndex        =   57
            Top             =   270
            Width           =   1365
            Begin VB.Label Label12 
               BackColor       =   &H0000FFFF&
               Caption         =   "添加(&T)"
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   45
               TabIndex        =   58
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
            TabIndex        =   55
            Top             =   540
            Width           =   1365
            Begin VB.Label Label11 
               BackColor       =   &H0000FFFF&
               Caption         =   "删除(&D)"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   45
               TabIndex        =   56
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
            TabIndex        =   53
            Top             =   1125
            Width           =   1365
            Begin VB.Label Label10 
               BackColor       =   &H0000FFFF&
               Caption         =   "清除(&L)"
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   45
               TabIndex        =   54
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
            TabIndex        =   51
            Top             =   855
            Width           =   1365
            Begin VB.Label Label4 
               BackColor       =   &H0000FFFF&
               Caption         =   "全选(&A)"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   45
               TabIndex        =   52
               Top             =   45
               Width           =   1635
            End
         End
         Begin VB.Label Label7 
            BackColor       =   &H0000FFFF&
            Caption         =   "编辑(&E)"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   90
            TabIndex        =   59
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
         TabIndex        =   38
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
            TabIndex        =   47
            Top             =   540
            Width           =   1365
            Begin VB.Label Label18 
               BackColor       =   &H0000FFFF&
               Caption         =   "打开(&O)"
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   45
               TabIndex        =   48
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
            TabIndex        =   45
            Top             =   810
            Width           =   1365
            Begin VB.Label Label13 
               BackColor       =   &H0000FFFF&
               Caption         =   "保存(&S)"
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   45
               TabIndex        =   46
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
            TabIndex        =   43
            Top             =   1485
            Width           =   1365
            Begin VB.Label Label14 
               BackColor       =   &H0000FFFF&
               Caption         =   "退出(&C)"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   45
               TabIndex        =   44
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
            TabIndex        =   41
            Top             =   1125
            Width           =   1365
            Begin VB.Label Label8 
               BackColor       =   &H0000FFFF&
               Caption         =   "汇入列表(&G)"
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   45
               TabIndex        =   42
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
            TabIndex        =   39
            Top             =   270
            Width           =   1365
            Begin VB.Label Label9 
               BackColor       =   &H0000FFFF&
               Caption         =   "新建(&N)"
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   45
               TabIndex        =   40
               Top             =   45
               Width           =   1635
            End
         End
         Begin VB.Label Label16 
            BackColor       =   &H0000FFFF&
            Caption         =   "文件(&F)"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   90
            TabIndex        =   49
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
         TabIndex        =   37
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
         Pattern         =   $"Cd Player12.frx":181B6
         System          =   -1  'True
         TabIndex        =   36
         ToolTipText     =   "Click to select a file"
         Top             =   630
         Width           =   2535
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   2730
         ItemData        =   "Cd Player12.frx":1827A
         Left            =   1935
         List            =   "Cd Player12.frx":1827C
         OLEDropMode     =   1  'Manual
         TabIndex        =   23
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
         TabIndex        =   34
         Top             =   270
         Width           =   4470
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            Caption         =   "所有文件"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   3375
            TabIndex        =   71
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
            TabIndex        =   70
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
            TabIndex        =   35
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
         BackColor       =   &H80000018&
         Caption         =   "选中文件"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   945
         TabIndex        =   64
         Top             =   5085
         Width           =   960
      End
      Begin VB.Label Label24 
         BackColor       =   &H80000018&
         Caption         =   "候选文件"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   945
         TabIndex        =   63
         Top             =   4815
         Width           =   960
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000018&
         Caption         =   "文件夹"
         ForeColor       =   &H00FF0000&
         Height          =   600
         Left            =   45
         TabIndex        =   62
         Top             =   4815
         Width           =   870
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000018&
         Caption         =   "驱动器   |文件类型"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   45
         TabIndex        =   61
         Top             =   4590
         Width           =   1860
      End
      Begin VB.Label Label21 
         BackColor       =   &H80000018&
         Caption         =   "菜单 + 工具"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   45
         TabIndex        =   60
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
      Left            =   11250
      TabIndex        =   65
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
      TabIndex        =   73
      Top             =   90
      Width           =   6000
   End
   Begin VB.Frame Frame13 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame10"
      ForeColor       =   &H80000008&
      Height          =   5460
      Left            =   6570
      TabIndex        =   101
      Top             =   450
      Width           =   4515
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   3735
      Top             =   6840
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   5760
      Picture         =   "Cd Player12.frx":1827E
      Top             =   6480
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Label lblday 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   4680
      TabIndex        =   108
      Top             =   7605
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      Height          =   210
      Left            =   4770
      TabIndex        =   107
      Top             =   7965
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label45 
      Caption         =   "end"
      Height          =   330
      Left            =   3420
      TabIndex        =   103
      Top             =   8010
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label Label44 
      Caption         =   "ed"
      Height          =   240
      Left            =   3420
      TabIndex        =   102
      Top             =   8370
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label28 
      Caption         =   "file"
      Height          =   195
      Left            =   1845
      TabIndex        =   69
      Top             =   8325
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label27 
      Caption         =   "Label27"
      Height          =   240
      Left            =   1980
      TabIndex        =   68
      Top             =   8010
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label26 
      Caption         =   "unlock"
      Height          =   240
      Left            =   2025
      TabIndex        =   67
      Top             =   7740
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   195
      Left            =   0
      TabIndex        =   17
      Top             =   8325
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   195
      Left            =   45
      TabIndex        =   3
      Top             =   8055
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   7785
      Visible         =   0   'False
      Width           =   1950
   End
End
Attribute VB_Name = "Form102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'RA.HideShowStatistics
'RA.SetPosition 10000
Option Explicit
Dim RmStop As Boolean
Dim RmGn As Boolean
Dim merlin As IAgentCtlCharacterEx
Const DATAPATH = "merlin.acs"

Private Declare Function RasHangUp Lib "RasApi32.DLL" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
Private Declare Function RasEnumConnections Lib "RasApi32.DLL" Alias "RasEnumConnectionsA" (lprasconn As Any, lpcb As Long, lpcConnections As Long) As Long
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










Dim AloT As Integer


Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' SetWindowPos Flags
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2

' SetWindowPos() hwndInsertAfter values
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Dim SearchFlag, Cdno As Integer
Dim co As Boolean
Public f200 As Integer
Public jd As Integer
Public jn, Skin As String
Dim selectedate%
Dim pIda As Long, pHnd As Long
Const SYNCHRONIZE = &H100000
Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Const INFINITE = &HFFFFFFFF
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Dim rgt As String
Dim hLB&, FileSpec$, UseFileSpec%
Dim TotalDirs%, TotalFiles%, Running%
Dim WFD As WIN32_FIND_DATA, hItem&, hFile&
Const vbBackslash = "\"
Const vbAllFiles = "*.*"
Const vbKeyDot = 46
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
Const SPI_SETSCREENSAVEACTIVE = 17
Const SPI_SETSCREENSAVETIMEOUT = 15
Const SPIF_SENDWININICHANGE = &H2
Const SPIF_UPDATEINIFILE = &H1

Private Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
     ByVal lpvParam As Long, ByVal fuWinIni As Long) As Long

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

Sub AutoPlay()
On Error Resume Next
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
      If FileExists(Label27.Caption + "Track01.cda") = True Then
      Update
        SendMCIString "play cd", True
        Playing = True
        CooLine1.Display = "CD Audio Mode  - Snowman Media  3.0"
Label1.Caption = "cd"
Frame1.Left = 45
Frame8.Left = 10000
TrackSelection.Left = 2880
  Exit Sub
         End If
    If FileExists(Label27.Caption + "MPEGAV\") = True Then
          File1.Path = Label27.Caption + "MPEGAV\"
                 File1.Pattern = "*.au;*.dat;*.and;*.aif;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.m3u;*.wma;*.wav;*.avi;*.mid;*.bmp;*.jpg;*.did;*.wmf;*.gif;*.rle;*.cur;*.emf"
          Dim ind As Integer
              For ind = 0 To File1.ListCount - 1
             ListFile.AddItem Label27.Caption + "MPEGAV\" + File1.List(ind), ListFile.ListCount
                  Next ind
            pid = ListFile.ListCount - File1.ListCount
                    MediaPlayer1.Filename = ListFile.List(pid)
         CooLine1.Display = "[ VCD Video ] MediaPlayer1.Filename  - Snowman Media  3.0"
        Label1.Caption = "media"
        TrackSelection.Left = 10000
       Frame8.Left = 2880
        Frame1.Left = 10000
        Exit Sub
        End If

    If Playing = False And Len(MediaPlayer1.Filename) = 0 Then
                 Dim firstpath As String, dircount As Integer, NumFiles As Integer
                  Dim Result As Long
   Dir1.Path = Label27.Caption
               File1.Pattern = "*.au;*.and;*.aif;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.m3u;*.wma;*.wav;*.avi;*.mid"
               If File1.Path = Label3.Caption Then Exit Sub
                 firstpath = Dir1.Path
                dircount = Dir1.ListCount
                 NumFiles = 0                       ' Reset global foundfiles indicator.
               Result = DirDiver(firstpath, dircount, "")
                  File1.Path = Dir1.Path
            
                    pid = ListFile.ListCount - Cdno
              MediaPlayer1.Filename = ListFile.List(pid)
            Cdno = 0
                Label1.Caption = "media"
        TrackSelection.Left = 10000
       Frame8.Left = 2880
        Frame1.Left = 10000
   Exit Sub
   End If
End Sub
Sub AutoPlayB()
 
     If Len(LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Name")) > 0 Then
             pid = -1
       MediaPlayer1.Filename = LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Name")
       MediaPlayer1.CurrentPosition = LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Rute")
      Dim idc As Integer
            For idc = 0 To ListFile.ListCount - 1
            If ListFile.List(idc) = MediaPlayer1.Filename Then
            pid = idc
            ListFile.ListIndex = pid
            Exit Sub
           End If
            Next
          If pid = -1 Then
          ListFile.AddItem (MediaPlayer1.Filename)
          pid = ListFile.ListCount - 1
          ListFile.ListIndex = pid
          End If
           Label1.Caption = "media"
        TrackSelection.Left = 10000
      Frame8.Left = 2880
        Frame1.Left = 10000
    End If
    Exit Sub

   
   


 
   
   
   
   
End Sub
Private Sub SearchDirs(curpath$)  ' curpath$ is passed w/ trailing "\"
    Dim dirs%, dirbuf$(), i%
    Label5.Caption = ""
    Label5.Caption = "正在搜索: " & curpath$
    DoEvents
    If Label44.Caption = "ed" Then Exit Sub
    If Not Running% Then Exit Sub
    hItem& = FindFirstFile(curpath$ & vbAllFiles, WFD)
    If hItem& <> INVALID_HANDLE_VALUE Then
        Do
            If (WFD.dwFileAttributes And vbDirectory) Then
                If Asc(WFD.cFileName) <> vbKeyDot Then
                    TotalDirs% = TotalDirs% + 1
                    If (dirs% Mod 10) = 0 Then ReDim Preserve dirbuf$(dirs% + 10)
                    dirs% = dirs% + 1
                    dirbuf$(dirs%) = Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
                End If
            ElseIf Not UseFileSpec% Then
                TotalFiles% = TotalFiles% + 1
            End If
        Loop While FindNextFile(hItem&, WFD)
        Call FindClose(hItem&)
    End If
    If UseFileSpec% Then
        SendMessage hLB&, WM_SETREDRAW, 0, 0
        Call SearchFileSpec(curpath$)
        SendMessage hLB&, WM_VSCROLL, SB_BOTTOM, 0
        SendMessage hLB&, WM_SETREDRAW, 1, 0
    End If
    For i% = 1 To dirs%: SearchDirs curpath$ & dirbuf$(i%) & vbBackslash: Next i%
End Sub
Private Sub SearchFileSpec(curpath$)   ' curpath$ is passed w/ trailing "\"
    hFile& = FindFirstFile(curpath$ & FileSpec$, WFD)
    If hFile& <> INVALID_HANDLE_VALUE Then
        Do
            DoEvents
            If Label44.Caption = "ed" Then Exit Sub
            If Not Running% Then Exit Sub
            SendMessage hLB&, LB_ADDSTRING, 0, _
                ByVal curpath$ & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
        Loop While FindNextFile(hFile&, WFD)
        Call FindClose(hFile&)
    End If
End Sub
Sub Find()
               List2.Clear
               Dim xh As Integer
          ScaleMode = vbPixels
          hLB& = List2.hwnd
          SendMessage hLB&, LB_INITSTORAGE, 30000&, ByVal 30000& * 200
         If Label44.Caption = "ed" Then Exit Sub
         ' If Running% Then: Running% = False: Exit Sub
    Dim drvbitmask&, maxpwr%, pwr%
    On Error Resume Next
    For xh = 1 To 13
    If xh = 1 Then FileSpec$ = "*.wa?"
    If xh = 2 Then FileSpec$ = "*.mp*"
    If xh = 3 Then FileSpec$ = "*.qt"
    If xh = 4 Then FileSpec$ = "*.wma"
    If xh = 5 Then FileSpec$ = "*.avi"
    If xh = 6 Then FileSpec$ = "*.mid"
    If xh = 7 Then FileSpec$ = "*.au"
    If xh = 8 Then FileSpec$ = "*.and"
    If xh = 9 Then FileSpec$ = "*.aif*"
    If xh = 10 Then FileSpec$ = "*.rmi"
    If xh = 11 Then FileSpec$ = "*.asx"
    If xh = 12 Then FileSpec$ = "*.m?v"
    If xh = 13 Then FileSpec$ = "*.asf"
    If MousePointer <> 11 Then MousePointer = 11
    Running% = True
    UseFileSpec% = True
    drvbitmask& = GetLogicalDrives()
        If drvbitmask& Then
            maxpwr% = Int(Log(drvbitmask&) / Log(2))   ' a little math...
        For pwr% = 0 To maxpwr%
            If Running% And (2 ^ pwr% And drvbitmask&) Then
            If Chr$(vbKeyA + pwr%) <> "A" And Chr$(vbKeyA + pwr%) <> "B" Then
                Call SearchDirs(Chr$(vbKeyA + pwr%) & ":\")
        End If
        End If
        Next
    End If
    Running% = False
    UseFileSpec% = False
   MousePointer = 0
   Label5.Caption = ""
   Label5.Caption = "共有媒体曲目: " & List2.ListCount
     Next
    Dim i As Integer
     Open Label3.Caption + "\SmMDb.dat" For Output As #1
    For i = 0 To List2.ListCount - 1
     Print #1, List2.List(i)
    Next i
   Close (1)
   MsgBox ("完成搜索,共发现媒体曲目: " & List2.ListCount & " 首.")
   Dim rgps As String
   rgps = InputBox("请输入本次更新媒体库的有关备注:")
    selectedate% = CInt(Format$(Now, "dd"))
Call fillcbomonth
Call fillcboyear
Call setdate
  LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "MediaWH_PS", rgps
  LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "MediaWH_TM", rgt
  LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "MediaWH_DA", lbldate.Caption
  LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "MediaWH_DY", lblday.Caption

   
   
   
   
   
   
   
   MsgBox ("初始化有关数据需要重新启动 Sowman Media  3.0;当你按下确定后 Snowman Media  3.0 将自动重新启动.")
   Form3.Show
    End Sub
Sub SetAlo()
'On Error Resume Next
LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Name", MediaPlayer1.Filename
LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Rute", MediaPlayer1.CurrentPosition
End Sub
Function StripNulls(startStrg$) As String
  Dim c%, item$
  c% = 1
  Do
    If Mid$(startStrg$, c%, 1) = Chr$(0) Then
      item$ = Mid$(startStrg$, 1, c% - 1)
      startStrg$ = Mid$(startStrg$, c% + 1, Len(startStrg$))
      StripNulls$ = item$
      Exit Function
    End If
    c% = c% + 1
  Loop
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
Function FileExists(Filename As String) As Boolean
On Error Resume Next
FileExists = Dir$(Filename) <> ""
If Err.Number <> 0 Then
FileExists = False
End If
On Error GoTo 0
End Function

Private Sub CDN_Arrival(ByVal Drive As String)
If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(14)) = 1 Then Call AutoPlay

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then Check2.Value = 0
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then Check1.Value = 0
End Sub























Private Sub Command1_Click()
Me.Caption = MediaPlayer1.ChannelDescription
End Sub

Private Sub Form_Resize()
If co = True Then
      If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Op_" + Str(2)) = True Then Call Tiday
      Dim i As Integer
         If Len(MediaPlayer1.Filename) > 0 Then
             pid = -1
                For i = 0 To ListFile.ListCount - 1
                      If ListFile.List(i) = MediaPlayer1.Filename Then
                         pid = i
                         ListFile.ListIndex = pid
            
                      End If
                Next
                      If pid = -1 Then
                           ListFile.AddItem (MediaPlayer1.Filename)
                            pid = ListFile.ListCount - 1
                            ListFile.ListIndex = pid
                            End If
         CooLine1.Display = "[" + Str(pid + 1) + " -" + Str(ListFile.ListCount) + " ] " + MediaPlayer1.Filename & "  - Snowman Media  3.0"
         Label1.Caption = "media"
         TrackSelection.Left = 10000
         Frame8.Left = 2880
         Frame1.Left = 10000
            Else
          If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(14)) = 1 Then
          Call AutoPlay
          If Playing = False And Len(MediaPlayer1.Filename) = 0 Then Call AutoPlayB
          If Playing = False And Len(MediaPlayer1.Filename) = 0 And ListFile.ListCount > 0 Then
                    pid = 0
                    MediaPlayer1.Filename = ListFile.List(ListFile.ListIndex)
          End If
          If LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Name") = MediaPlayer1.Filename Then
          MediaPlayer1.CurrentPosition = LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Rute")
          End If
          If Playing = False And Len(MediaPlayer1.Filename) = 0 Then MediaPlayer1.Filename = Label3.Caption + "\SmM_St.wma"

          End If
          End If
        co = False


End If
If Me.WindowState = 0 Then
On Error Resume Next
Me.Height = myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "h", "")
Me.Width = myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "w", "")
End If
'Label47.Caption = Str(Me.Height) + "   " + Str(Me.Width)
End Sub

Private Sub Frame11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call reco2
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
If Len(File1.Filename) > 0 Then
     If Len(Dir1.Path) = 3 Then
     List1.AddItem Dir1.Path & File1.Filename
     Else: List1.AddItem Dir1.Path & "\" & File1.Filename
      End If
     Else: MsgBox ("还没选定要添加的曲目,请先选定好要添加的曲目再添加入列表.")
    
End If
Label28.Caption = "file"
End Sub
Private Sub Label13_Click()
If List1.ListCount > 0 Then
  CommonDialog1.Filter = "列表文件:M3u" & _
          "|*.M3u|所有文件:*.*|*.*"
CommonDialog1.Filename = ""
CommonDialog1.ShowSave
If Len(CommonDialog1.Filename) > 0 Then
Dim i As Integer
   Open CommonDialog1.Filename For Output As #1
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
          "|*.M3u|所有文件:*.*|*.*"
CommonDialog1.Filename = ""
CommonDialog1.ShowSave
If Len(CommonDialog1.Filename) > 0 Then
Dim i As Integer
   Open CommonDialog1.Filename For Output As #1
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
          "|*.M3u|所有文件:*.*|*.*"
CommonDialog1.Filename = ""
CommonDialog1.ShowSave
If Len(CommonDialog1.Filename) > 0 Then
Dim i As Integer
   Open CommonDialog1.Filename For Output As #1
    For i = 0 To List1.ListCount - 1
     Print #1, List1.List(i)
    Next i
   Close (1)
End If
End If
   CommonDialog1.Filter = "列表文件:M3u" & _
      "|*.M3u|所有文件:*.*|*.*"
CommonDialog1.Filename = ""
CommonDialog1.ShowOpen
If Len(CommonDialog1.Filename) > 0 Then
List1.Clear
Dim test As String
 Open CommonDialog1.Filename For Input As #1
    While Not EOF(1)
    Line Input #1, test
    List1.AddItem RTrim(test)
    Wend
    Close #1
End If
End Sub

Private Sub Label29_Click()
Label44.Caption = "st"
Call Find
End Sub
Private Sub Label29_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture19.BackColor = &HFF0000
Label29.BackColor = &HFF0000
Label29.ForeColor = &HFFFF&
Picture20.BackColor = &HFFFF&
Label30.BackColor = &HFFFF&
Label30.ForeColor = &HFF0000
Picture21.BackColor = &HFFFF&
Label31.BackColor = &HFFFF&
Label31.ForeColor = &HFF0000
Picture26.BackColor = &HFFFF&
Label43.BackColor = &HFFFF&
Label43.ForeColor = &HFF0000
Picture25.BackColor = &HFFFF&
Label42.BackColor = &HFFFF&
Label42.ForeColor = &HFF0000
Picture24.BackColor = &HFFFF&
Label41.BackColor = &HFFFF&
Label41.ForeColor = &HFF0000
Picture22.BackColor = &HFFFF&
Label40.BackColor = &HFFFF&
Label40.ForeColor = &HFF0000
End Sub
Private Sub Label30_Click()
If Label44.Caption = "st" Then
MsgBox ("为了更好地使用本功能,请另找时间完成媒体库信息的更新.")
Label44.Caption = "ed"
End If
End Sub

Private Sub Label30_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture20.BackColor = &HFF0000
Label30.BackColor = &HFF0000
Label30.ForeColor = &HFFFF&
Picture21.BackColor = &HFFFF&
Label31.BackColor = &HFFFF&
Label31.ForeColor = &HFF0000
Picture26.BackColor = &HFFFF&
Label43.BackColor = &HFFFF&
Label43.ForeColor = &HFF0000
Picture25.BackColor = &HFFFF&
Label42.BackColor = &HFFFF&
Label42.ForeColor = &HFF0000
Picture24.BackColor = &HFFFF&
Label41.BackColor = &HFFFF&
Label41.ForeColor = &HFF0000
Picture22.BackColor = &HFFFF&
Label40.BackColor = &HFFFF&
Label40.ForeColor = &HFF0000
Picture19.BackColor = &HFFFF&
Label29.BackColor = &HFFFF&
Label29.ForeColor = &HFF0000
End Sub

Private Sub Label31_Click()
If List2.ListCount > 0 Then
Dim i As Integer
    For i = 0 To List2.ListCount - 1
     ListFile.AddItem List2.List(i)
    Next i

 pid = ListFile.ListCount - List2.ListCount
                    MediaPlayer1.Filename = ListFile.List(pid)
Else: MsgBox ("媒体库信息为空,无法播放.请先更新媒体库信息.")
End If
End Sub

Private Sub Label31_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture19.BackColor = &HFFFF&
Label29.BackColor = &HFFFF&
Label29.ForeColor = &HFF0000
Picture20.BackColor = &HFFFF&
Label30.BackColor = &HFFFF&
Label30.ForeColor = &HFF0000
Picture21.BackColor = &HFF0000
Label31.BackColor = &HFF0000
Label31.ForeColor = &HFFFF&
Picture26.BackColor = &HFFFF&
Label43.BackColor = &HFFFF&
Label43.ForeColor = &HFF0000
Picture25.BackColor = &HFFFF&
Label42.BackColor = &HFFFF&
Label42.ForeColor = &HFF0000
Picture24.BackColor = &HFFFF&
Label41.BackColor = &HFFFF&
Label41.ForeColor = &HFF0000
Picture22.BackColor = &HFFFF&
Label40.BackColor = &HFFFF&
Label40.ForeColor = &HFF0000



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

Private Sub Label40_Click()
 CommonDialog1.Filter = "媒体文件:Mp3、Wma、Wav、Wax、Asf、Rmi、Asx、Mov、M1v、Mp2、Mpg、Mpeg、Mpa、Mpe、Avi、Mid、Qt、M3u、Aif、Aifc、Aiff、Au、Snd..." & _
          "|*.au;*.and;*.aif;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.m3u;*.wma;*.wav;*.avi;*.mid|所有文件:*.*|*.*"
          CommonDialog1.FilterIndex = 1
          CommonDialog1.Filename = ""
          CommonDialog1.ShowOpen
          If Len(CommonDialog1.Filename) > 0 Then
         List2.AddItem CommonDialog1.Filename
            Dim i As Integer
     Open Label3.Caption + "\SmMDb.dat" For Output As #1
    For i = 0 To List2.ListCount - 1
     Print #1, List2.List(i)
    Next i
   Close (1)
          End If
End Sub

Private Sub Label40_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture19.BackColor = &HFFFF&
Label29.BackColor = &HFFFF&
Label29.ForeColor = &HFF0000
Picture20.BackColor = &HFFFF&
Label30.BackColor = &HFFFF&
Label30.ForeColor = &HFF0000
Picture21.BackColor = &HFFFF&
Label31.BackColor = &HFFFF&
Label31.ForeColor = &HFF0000
Picture26.BackColor = &HFFFF&
Label43.BackColor = &HFFFF&
Label43.ForeColor = &HFF0000
Picture25.BackColor = &HFFFF&
Label42.BackColor = &HFFFF&
Label42.ForeColor = &HFF0000
Picture24.BackColor = &HFFFF&
Label41.BackColor = &HFFFF&
Label41.ForeColor = &HFF0000
Picture22.BackColor = &HFF0000
Label40.BackColor = &HFF0000
Label40.ForeColor = &HFFFF&



End Sub

Private Sub Label41_Click()
If Len(List2.List(List2.ListIndex)) > 0 Then
If List2.ListCount > 0 And List2.SelCount > 0 Then
List2.RemoveItem (List2.ListIndex)
  Dim i As Integer
     Open Label3.Caption + "\SmMDb.dat" For Output As #1
    For i = 0 To List2.ListCount - 1
     Print #1, List2.List(i)
    Next i
   Close (1)
End If
Else: MsgBox ("还没选定要删除的曲目,请先选定好要删除的曲目再从库中删除.")
End If
End Sub
Private Sub Label41_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture19.BackColor = &HFFFF&
Label29.BackColor = &HFFFF&
Label29.ForeColor = &HFF0000
Picture20.BackColor = &HFFFF&
Label30.BackColor = &HFFFF&
Label30.ForeColor = &HFF0000
Picture21.BackColor = &HFFFF&
Label31.BackColor = &HFFFF&
Label31.ForeColor = &HFF0000
Picture26.BackColor = &HFFFF&
Label43.BackColor = &HFFFF&
Label43.ForeColor = &HFF0000
Picture25.BackColor = &HFFFF&
Label42.BackColor = &HFFFF&
Label42.ForeColor = &HFF0000
Picture24.BackColor = &HFF0000
Label41.BackColor = &HFF0000
Label41.ForeColor = &HFFFF&
Picture22.BackColor = &HFFFF&
Label40.BackColor = &HFFFF&
Label40.ForeColor = &HFF0000
End Sub

Private Sub Label42_Click()
  CommonDialog1.Filename = ""
               CommonDialog1.Filter = "列表文件:M3u" & _
          "|*.M3u|所有文件:*.*|*.*"
             CommonDialog1.ShowOpen
             If Len(CommonDialog1.Filename) > 0 Then
              Open CommonDialog1.Filename For Input As #1
           While Not EOF(1)
          Line Input #1, test
           List2.AddItem RTrim(test)
           Wend
             Close #1
               Dim i As Integer
     Open Label3.Caption + "\SmMDb.dat" For Output As #1
    For i = 0 To List2.ListCount - 1
     Print #1, List2.List(i)
    Next i
   Close (1)
            End If
End Sub

Private Sub Label42_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture19.BackColor = &HFFFF&
Label29.BackColor = &HFFFF&
Label29.ForeColor = &HFF0000
Picture20.BackColor = &HFFFF&
Label30.BackColor = &HFFFF&
Label30.ForeColor = &HFF0000
Picture21.BackColor = &HFFFF&
Label31.BackColor = &HFFFF&
Label31.ForeColor = &HFF0000
Picture26.BackColor = &HFFFF&
Label43.BackColor = &HFFFF&
Label43.ForeColor = &HFF0000
Picture25.BackColor = &HFF0000
Label42.BackColor = &HFF0000
Label42.ForeColor = &HFFFF&
Picture24.BackColor = &HFFFF&
Label41.BackColor = &HFFFF&
Label41.ForeColor = &HFF0000
Picture22.BackColor = &HFFFF&
Label40.BackColor = &HFFFF&
Label40.ForeColor = &HFF0000
End Sub

Private Sub Label43_Click()
            If List2.ListCount > 0 Then
                  Dim i As Integer
              CommonDialog1.Filename = ""
               CommonDialog1.Filter = "列表文件:M3u" & _
          "|*.M3u|所有文件:*.*|*.*"
             CommonDialog1.ShowSave
             If Len(CommonDialog1.Filename) > 0 Then
             Open CommonDialog1.Filename For Output As #1
    For i = 0 To List2.ListCount - 1
     Print #1, List2.List(i)
    Next i
   Close (1)
   End If
        Else: MsgBox ("当前媒体库为空无法保存,请先更新媒体库信息.")
        End If
End Sub

Private Sub Label43_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Picture19.BackColor = &HFFFF&
Label29.BackColor = &HFFFF&
Label29.ForeColor = &HFF0000
Picture20.BackColor = &HFFFF&
Label30.BackColor = &HFFFF&
Label30.ForeColor = &HFF0000
Picture21.BackColor = &HFFFF&
Label31.BackColor = &HFFFF&
Label31.ForeColor = &HFF0000
Picture26.BackColor = &HFF0000
Label43.BackColor = &HFF0000
Label43.ForeColor = &HFFFF&
Picture25.BackColor = &HFFFF&
Label42.BackColor = &HFFFF&
Label42.ForeColor = &HFF0000
Picture24.BackColor = &HFFFF&
Label41.BackColor = &HFFFF&
Label41.ForeColor = &HFF0000
Picture22.BackColor = &HFFFF&
Label40.BackColor = &HFFFF&
Label40.ForeColor = &HFF0000
End Sub



Private Sub Label8_Click()
   CommonDialog1.Filename = ""
               CommonDialog1.Filter = "列表文件:M3u" & _
          "|*.M3u|所有文件:*.*|*.*"
             CommonDialog1.ShowOpen
             If Len(CommonDialog1.Filename) > 0 Then
              Open CommonDialog1.Filename For Input As #1
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
          "|*.M3u|所有文件:*.*|*.*"
CommonDialog1.Filename = ""
CommonDialog1.ShowSave
If Len(CommonDialog1.Filename) > 0 Then
Dim i As Integer
   Open CommonDialog1.Filename For Output As #1
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

Private Sub List2_DblClick()
 If Len(List2.List(List2.ListIndex)) > 0 Then
                  ListFile.AddItem List2.List(List2.ListIndex)
                pid = ListFile.ListCount - 1
                    MediaPlayer1.Filename = ListFile.List(pid)
          End If
End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call reco2
End Sub

Private Sub List2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim ThisFile As Variant
    For Each ThisFile In Data.Files
        List2.AddItem ThisFile
    Next
End Sub

Private Sub ListFile_Click()
ListFile.ToolTipText = "[" + Str(ListFile.ListIndex + 1) + " -" + Str(ListFile.ListCount) + " ] " + ListFile.List(ListFile.ListIndex) + "  - Snowman Media  3.0"
End Sub
Private Sub MediaPlayer1_Click(Button As Integer, ShiftState As Integer, x As Single, y As Single)
On Error Resume Next
Label1.Caption = "media"
Frame1.Left = 10000
TrackSelection.Left = 10000
 Frame8.Left = 2880
If Button = 1 Then MediaPlayer1.Play
If Button = 2 Then MediaPlayer1.Pause
End Sub
Private Sub MediaPlayer1_DblClick(Button As Integer, ShiftState As Integer, x As Single, y As Single)
On Error Resume Next
Label1.Caption = "media"
Frame1.Left = 10000
TrackSelection.Left = 10000
 Frame8.Left = 2880
 MediaPlayer1.Filename = "LLXX"
If Button = 1 Then Call Image4_Click
If Button = 2 Then Call Image5_Click
MediaPlayer1.Play
End Sub
Private Sub MediaPlayer1_Error()
Dim T As String
'If MediaPlayer1.ErrorCode = -2147220891 Then
    T = UCase(Right(ListFile.List(pid), 3))
     If T = ".RA" Or T = ".RM" Or T = "RAM" Or T = ".RT" Or T = ".RP" Or T = "SMI" Or T = "MIL" Then
        RA.Source = ListFile.List(pid)
        'RA.DoPlay
     End If
'End If
End Sub
Private Sub Option1_Click()
File1.Pattern = "*.au;*.and;*.aif;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.m3u;*.wma;*.wav;*.avi;*.mid"
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
List1.AddItem Dir1.Path & File1.Filename
Else: List1.AddItem Dir1.Path & "\" & File1.Filename
End If
End If
Label28.Caption = "file"
End Sub
Private Sub File1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 2 Then Exit Sub
        If Len(File1.Filename) > 0 Then
              If Len(Dir1.Path) = 3 Then
                    ListFile.AddItem Dir1.Path & File1.Filename
                    pid = ListFile.ListCount - 1
                    MediaPlayer1.Filename = ListFile.List(pid)
                    Else:
                    ListFile.AddItem Dir1.Path & "\" & File1.Filename
                    pid = ListFile.ListCount - 1
                    MediaPlayer1.Filename = ListFile.List(pid)
              End If
        End If
Label28.Caption = "file"
End Sub
Private Sub File1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call reco
Label28.Caption = "file"
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
                  Frame11.Left = 10000
               Frame13.Left = 10000
 End If
End Sub
Sub ToolB()
      Frame2.Left = 10000
            Frame5.Left = 10000
             Frame6.Left = 10000
             Frame7.Left = 10000
              Frame9.Left = 1530
               Frame10.Left = 1575
                  List1.Clear
                  Frame11.Left = 10000
               Frame13.Left = 10000
End Sub
Sub ToolC()
      Frame2.Left = 10000
            Frame5.Left = 10000
             Frame6.Left = 10000
             Frame7.Left = 10000
              Frame9.Left = 10000
               Frame10.Left = 10000
                  List1.Clear
                  Frame11.Left = 1530
               Frame13.Left = 1575
End Sub

'Private Sub WB_DocumentComplete(ByVal pDisp As Object, url As Variant)
'If WB.Left = 0 Then CooLine1.Display = "[ H2ont Media Channer ] " + url + "  - Snowman Media  3.0"
'End Sub


Private Sub Form_Load()
 
If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(12)) = 1 Then MediaPlayer1.ClickToPlay = True
If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(7)) = 1 Then Call DiPb
If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(8)) = 1 Then SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(11)) = 0 Then ctListBar1.SmoothScroll = False
ctListBar1.BackImage = LoadPicture(LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Te_" + Str(2)))
If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Op_" + Str(15)) = True Then
CooLine1.DisplayStyle = 4
CooLine1.InsCharacter = " "
End If
If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Op_" + Str(16)) = True Then
CooLine1.DisplayStyle = 2
CooLine1.InsCharacter = " "
End If
If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Op_" + Str(17)) = True Then
CooLine1.DisplayStyle = 4
CooLine1.InsCharacter = ""
End If
If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(1)) = 0 Then ListFile.OLEDropMode = 0




Label1.Caption = "media"
Frame1.Left = 10000
TrackSelection.Left = 10000
 Frame8.Left = 2880
Label3.Caption = App.Path

Skin = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Skin_Path")
  SkinForm1.SkinPath = Skin
      co = True
    ctListBar1.AddList "当前列表"
    ctListBar1.AddList "媒体播放"
    ctListBar1.AddList "在线播放"
    ctListBar1.AddList "媒体书签"
    ctListBar1.AddList "其他功能"
    ctListBar1.AddList "关于我们"
    ctListBar1.AddListImage 1, "Snowflake", 21
    ctListBar1.AddListImage 1, "视频窗口", 25
    ctListBar1.AddListImage 1, "全屏欣赏", 26
    ctListBar1.AddListImage 1, "锁定播放", 35
    ctListBar1.AddListImage 1, "标记书签", 41
    'ctListBar1.AddListImage 1, "媒体说明", 33
    'ctListBar1.AddListImage 1, "媒体属性", 5
    'ctListBar1.AddListImage 1, "统计信息", 6
    
    ctListBar1.AddListImage 2, "添加曲目", 1
    ctListBar1.AddListImage 2, "添加 URL", 44
    ctListBar1.AddListImage 2, "删除曲目", 24
    ctListBar1.AddListImage 2, "添加目录", 30
    ctListBar1.AddListImage 2, "汇入列表", 36
    ctListBar1.AddListImage 2, "智能整理", 3
    ctListBar1.AddListImage 2, "导出保存", 4
    ctListBar1.AddListImage 2, "清空列表", 2
    
    'ctListBar1.AddListImage 3, "媒体播放", 7
    ctListBar1.AddListImage 3, "VCD 播放", 32
    ctListBar1.AddListImage 3, "DVD 播放", 31
    ctListBar1.AddListImage 3, "媒体光盘播放", 9
    ctListBar1.AddListImage 3, "Flash 播放", 8
    ctListBar1.AddListImage 3, "3DS 播放", 10
    
    ctListBar1.AddListImage 4, "在线媒体", 11
    ctListBar1.AddListImage 4, "在线 Flash", 12
    
    ctListBar1.AddListImage 5, "断点续播", 19
    ctListBar1.AddListImage 5, "媒体书签 [A]", 39
    ctListBar1.AddListImage 5, "媒体书签 [B]", 39
    ctListBar1.AddListImage 5, "媒体书签 [C]", 39
    ctListBar1.AddListImage 5, "媒体书签 [D]", 39
    'ctListBar1.AddListImage 4, "媒体指南", 45
   

    'ctListBar1.AddListImage 5, "列表播放", 16
    ctListBar1.AddListImage 5, "编辑列表", 17
    'ctListBar1.AddListImage 5, "目录播放", 40
    ctListBar1.AddListImage 5, "媒体库", 18

    ctListBar1.AddListImage 6, "数字音频", 22
    ctListBar1.AddListImage 6, "图片浏览", 34
    ctListBar1.AddListImage 6, "个性化播放", 43
    ctListBar1.AddListImage 6, "自动关机", 28
    ctListBar1.AddListImage 6, "设置选项", 42
    ctListBar1.AddListImage 7, "帮助主题", 50
    ctListBar1.AddListImage 7, "显示助手", 46
    ctListBar1.AddListImage 7, "检查更新", 47
    ctListBar1.AddListImage 7, "关于 Sm.M.", 48
    ctListBar1.AddListImage 7, "联系我们", 49
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
If (App.PrevInstance = True) Then
    End
End If
'Timer1.Enabled = False
FastForwardSpeed = 10
CDLoad = False
If (SendMCIString("open cdaudio alias cd wait shareable", True) = False) Then
    End
End If
SendMCIString "set cd time format tmsf wait", True
Timer1.Enabled = True
   Open Label3.Caption + "\SmM30.dat" For Input As #1
    While Not EOF(1)
    Line Input #1, test
    ListFile.AddItem RTrim(test)
    Wend
    Close #1

 RA.SetNoLogo True
  RA.SetEnableContextMenu False
Form103.Label1.Caption = "ED"

End Sub
Private Sub lblnumber_click(Index As Integer)
Dim i%
On Error GoTo err1
For i% = 0 To 30
Next i%
selectedate% = Index + 1
Dim month1%, day1%, year1%, date1$
day1% = selectedate%
month1% = cbomonth.ListIndex + 1
year1% = cboyear.ListIndex + 1960
date1$ = (Str$(month1%) + "/" + Str$(day1%) + "/" + Str$(year1%))
'date1$ = Format$(date1$, "general date")
Dim r%
Dim caption1$
r% = Weekday(date1$)
If r% = 1 Then
    caption1$ = "Sunday"
ElseIf r% = 2 Then
    caption1 = "Monday"
ElseIf r% = 3 Then
    caption1 = "Tuesday"
ElseIf r% = 4 Then
    caption1 = "Wednesday"
ElseIf r% = 5 Then
    caption1 = "Thursday"
ElseIf r% = 6 Then
    caption1 = "Friday"
Else
    caption1 = "Saturday"
End If
lblday.Caption = caption1$
lbldate.Caption = Format$(date1$, "long date")
err1:
    If Err = 0 Then Exit Sub
    If Err = 13 Then
        selectedate% = selectedate% - 1
    Exit Sub
    End If
    End Sub

Private Sub setdate()
Dim r%, i%
r% = CInt(Format$(Now, "yyyy"))
i% = r% - 1960
cboyear.ListIndex = i%
r% = CInt(Format$(Now, "mm"))
cbomonth.ListIndex = (r% - 1)
End Sub
Private Sub cbomonth_click()

Call lblnumber_click(selectedate% - 1)
End Sub
Private Sub fillcbomonth()
cbomonth.AddItem "January"
cbomonth.AddItem "February"
cbomonth.AddItem "March"
cbomonth.AddItem "April"
cbomonth.AddItem "May"
cbomonth.AddItem "June"
cbomonth.AddItem "July"
cbomonth.AddItem "August"
cbomonth.AddItem "September"
cbomonth.AddItem "October"
cbomonth.AddItem "November"
cbomonth.AddItem "December"
End Sub
Private Sub fillcboyear()
Dim i%
For i% = 1960 To 2060 'put whatever years tyou want here,
    cboyear.AddItem Str$(i%) 'but don't forget to also change the code in setdate
Next i%
End Sub
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim ThisFile As Variant
    For Each ThisFile In Data.Files
        ListFile.AddItem ThisFile
    Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(2)) = 1 Then ListFile.Clear
Call EnPb
If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Op_" + Str(3)) = True Then Call Tiday
If Len(MediaPlayer1.Filename) > 0 And LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Op_" + Str(6)) = True Then Call SetAlo
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
If Label45.Caption = "end" Then End
End Sub
Private Sub Image16_Click()
If ListFile.ListCount = 0 Then
MsgBox ("同步媒体列表为空,无可用的媒体用于播放.请添加媒体曲目到同步列表后再尝试启用媒体播放模式.")
Exit Sub
End If
CooLine1.Display = "Media Mode  - Snowman Media  3.0"
Label1.Caption = "media"
Frame1.Left = 10000
TrackSelection.Left = 10000
 Frame8.Left = 2880
End Sub
Private Sub Image17_Click()
If Len(TrackTime.Caption) > 0 Then
CooLine1.Display = "CD Audio Mode  - Snowman Media  3.0"
Label1.Caption = "cd"
Frame1.Left = 45
Frame8.Left = 10000
TrackSelection.Left = 2880
Else: MsgBox ("找不到 CD 唱片,无法启用 CD 播放模式.请插入 CD 唱片后重试.")
End If
End Sub
Private Sub image18_Click()
 On Error Resume Next
If Label1.Caption = "rm" Then
If RA.Left = 0 Then RmStop = False
RA.DoPlay
Exit Sub
End If
If Label1.Caption = "cd" Then
SendMCIString "play cd", True
Playing = True
If Len(TrackTime.Caption) > 0 And Playing = True Then
If RA.Left = 0 Then RmStop = False
If Len(MediaPlayer1.Filename) > 0 Then MediaPlayer1.Filename = "LLXX"
If RA.Left = 0 Then RmStop = False
RA.DoStop
RA.Left = 10000
SLD1.Left = 10000
End If
Exit Sub
End If
If Label1.Caption = "media" Then
         If Len(MediaPlayer1.Filename) = 0 Then
                       If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(14)) = 1 Then
                             Call AutoPlay
                             If Playing = False And Len(MediaPlayer1.Filename) = 0 Then Call AutoPlayB
                             If LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Name") = MediaPlayer1.Filename Then MediaPlayer1.CurrentPosition = LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Rute")
                       Else
                             If Playing = False And Len(MediaPlayer1.Filename) = 0 And ListFile.ListCount > 0 Then
                                          pid = 0
                                          MediaPlayer1.Filename = ListFile.List(pid)
                             End If
                       End If
                       
         End If
  MediaPlayer1.Play
End If

If f200 = 1 Then
  Form200.MediaPlayer1.Filename = "LLXX"
         If Len(Form200.MediaPlayer1.Filename) = 0 Then
                       If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(14)) = 1 Then
                             Call AutoPlay
                             If Playing = False And Len(Form200.MediaPlayer1.Filename) = 0 Then Call AutoPlayB
                             If LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Name") = Form200.MediaPlayer1.Filename Then Form200.MediaPlayer1.CurrentPosition = LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Rute")
                       Else
                             If Playing = False And Len(Form200.MediaPlayer1.Filename) = 0 And ListFile.ListCount > 0 Then
                                          pid = 0
                                          Form200.MediaPlayer1.Filename = ListFile.List(pid)
                             End If
                       End If
                       
         End If
  Form200.MediaPlayer1.Play
End If








End Sub
Private Sub Image3_Click()
If Label1.Caption = "rm" Then
RA.DoPause
Exit Sub
End If
If Label1.Caption = "cd" Then
SendMCIString "pause cd", True
Playing = False
Update
If Len(TrackTime.Caption) > 0 And Playing = True Then
If RA.Left = 0 Then RmStop = False
If Len(MediaPlayer1.Filename) > 0 Then MediaPlayer1.Filename = "LLXX"
If RA.Left = 0 Then RmStop = False
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
MediaPlayer1.Filename = "LLXX"
Form200.MediaPlayer1.Pause
End If







End If
End Sub
Private Sub image11_Click()
If Form102.Eject.Enabled = True Then
SendMCIString "stop cd wait", True
Command = "seek cd to " & Track
SendMCIString Command, True
Playing = False
SendMCIString "set cd door open", True
Else
 SendMCIString "set cd door closed", True
 End If
Update
End Sub
Private Sub Image7_Click()
On Error Resume Next

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
If RA.Left = 0 Then RmStop = False
If Len(MediaPlayer1.Filename) > 0 Then MediaPlayer1.Filename = "LLXX"
If RA.Left = 0 Then RmStop = False
RA.DoStop
RA.Left = 10000
SLD1.Left = 10000
End If
Exit Sub
End If
If Label1.Caption = "media" Then
MediaPlayer1.CurrentPosition = MediaPlayer1.CurrentPosition + 5




If f200 = 1 Then
MediaPlayer1.Filename = "LLXX"
Form200.MediaPlayer1.CurrentPosition = Form200.MediaPlayer1.CurrentPosition + 5
End If


End If
End Sub
Private Sub Image6_Click()
On Error Resume Next
If Label1.Caption = "rm" Then
If RA.GetPosition - 5 < 0 Then
RA.SetPosition -1

Else: RA.SetPosition RA.GetPosition - 5
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
If RA.Left = 0 Then RmStop = False
If Len(MediaPlayer1.Filename) > 0 Then MediaPlayer1.Filename = "LLXX"
If RA.Left = 0 Then RmStop = False
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
MediaPlayer1.Filename = "LLXX"
If Form200.MediaPlayer1.CurrentPosition - 5 < 0 Then
Form200.MediaPlayer1.CurrentPosition = -1

Else: Form200.MediaPlayer1.CurrentPosition = Form200.MediaPlayer1.CurrentPosition - 5
End If


End If
End If
End Sub
Public Sub Image4_Click()
On Error Resume Next
If RA.Left = 0 Then RmStop = False
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
If RA.Left = 0 Then RmStop = False
If Len(MediaPlayer1.Filename) > 0 Then MediaPlayer1.Filename = "LLXX"
If RA.Left = 0 Then RmStop = False
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
              MediaPlayer1.Filename = ListFile.List(pid)
     End If
     If f200 = 1 Then
             MediaPlayer1.Filename = "LLXX"
                 pid = pid + 1
                     If pid = ListFile.ListCount Then
                             pid = 0
                     End If
              Form200.MediaPlayer1.Filename = ListFile.List(pid)
      End If
End If
End Sub
Public Sub Image5_Click()
On Error Resume Next
If RA.Left = 0 Then RmStop = False
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
If RA.Left = 0 Then RmStop = False
If Len(MediaPlayer1.Filename) > 0 Then MediaPlayer1.Filename = "LLXX"
If RA.Left = 0 Then RmStop = False
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
                MediaPlayer1.Filename = ListFile.List(pid)
          End If
            If f200 = 1 Then
           MediaPlayer1.Filename = "LLXX"
           pid = pid - 1
           If pid < 0 Then pid = 0
        Form200.MediaPlayer1.Filename = ListFile.List(pid)
End If


End If
End Sub
Private Sub Image2_Click()
If Label1.Caption = "cd" Then
SendMCIString "stop cd wait", True
Command = "seek cd to " & Track
SendMCIString Command, True
Playing = False
Update
End If
If Label1.Caption = "media" Then
Call SetAlo
MediaPlayer1.Filename = "LLXX"


If f200 = 1 Then Form200.MediaPlayer1.Filename = "LLXX"







End If
End Sub
Private Function fSetVolumeControl(ByVal hmixer As Long, _
    mxc As MIXERCONTROL, ByVal Volume As Long) As Boolean
Dim rc   As Long
Dim mxcd As MIXERCONTROLDETAILS
Dim vol  As MIXERCONTROLDETAILS_UNSIGNED
With mxcd
    .item = 0
    .dwControlID = mxc.dwControlID
    .cbStruct = Len(mxcd)
    .cbDetails = Len(vol)
End With
hmem = GlobalAlloc(&H40, Len(vol))
mxcd.paDetails = GlobalLock(hmem)
mxcd.cChannels = 1
vol.dwValue = Volume
Call CopyPtrFromStruct(mxcd.paDetails, vol, Len(vol))
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
        If Len(File1.Filename) > 0 Then
              If Len(Dir1.Path) = 3 Then
                    ListFile.AddItem Dir1.Path & File1.Filename
                    pid = ListFile.ListCount - 1
                    MediaPlayer1.Filename = ListFile.List(pid)
                    Else:
                    ListFile.AddItem Dir1.Path & "\" & File1.Filename
                    pid = ListFile.ListCount - 1
                    MediaPlayer1.Filename = ListFile.List(pid)
              End If
            Else: MsgBox ("还没有选定要预览的曲目,请先选择好要预览的曲目再预览.")
        End If
End If
If Label28.Caption = "list" Then
         If Len(List1.List(List1.ListIndex)) > 0 Then
                  ListFile.AddItem List1.List(List1.ListIndex)
                pid = ListFile.ListCount - 1
                    MediaPlayer1.Filename = ListFile.List(pid)
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
MediaPlayer1.Filename = ListFile.List(ListFile.ListCount - ib)
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
                    MediaPlayer1.Filename = ListFile.List(pid)
          End If
Label28.Caption = "list"
End Sub
Private Sub ListFile_DblClick()
RmStop = False
pid = ListFile.ListIndex
MediaPlayer1.Filename = ListFile.List(pid)
 If f200 = 1 Then
 MediaPlayer1.Filename = "LLXX"
 Form200.MediaPlayer1.Filename = ListFile.List(pid)
 MediaPlayer1.Filename = "LLXX"
End If
Label1.Caption = "media"
Frame1.Left = 10000
TrackSelection.Left = 10000
 Frame8.Left = 2880
End Sub
Private Sub ListFile_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
If ListFile.ListCount > 0 And ListFile.SelCount > 0 Then
If pid <= ListFile.ListIndex Then pid = pid - 1
ListFile.RemoveItem (ListFile.ListIndex)
End If
End If
End Sub
Private Sub Listfile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim ThisFile As Variant
    For Each ThisFile In Data.Files
        ListFile.AddItem ThisFile
    Next
End Sub
Sub mpn()
If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(3)) = 1 Then
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
MediaPlayer1.Filename = MediaPlayer1.Filename
Exit Sub
End If
If Check1.Value = 0 And Check2.Value = 0 Then
pid = pid + 1
If pid = ListFile.ListCount Then
MediaPlayer1.Filename = "LLXX"
     Form4.Asd = True
End If
MediaPlayer1.Filename = ListFile.List(pid)
Exit Sub
End If
If Check1.Value = 1 Then
pid = pid + 1
If pid = ListFile.ListCount Then
pid = 0
End If
MediaPlayer1.Filename = ListFile.List(pid)
Exit Sub
End If
If Check2.Value = 1 Then
 Randomize
    pid = Int(ListFile.ListCount * Rnd)
MediaPlayer1.Filename = ListFile.List(pid)
Exit Sub
End If
End Sub
Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
     If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(3)) = 1 Then
 
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
MediaPlayer1.Filename = MediaPlayer1.Filename
Exit Sub
End If

If Check1.Value = 0 And Check2.Value = 0 Then
pid = pid + 1
If pid = ListFile.ListCount Then
MediaPlayer1.Filename = "LLXX"
     Form4.Asd = True
End If
MediaPlayer1.Filename = ListFile.List(pid)
Exit Sub
End If

If Check1.Value = 1 Then
pid = pid + 1
If pid = ListFile.ListCount Then
pid = 0
End If
MediaPlayer1.Filename = ListFile.List(pid)
Exit Sub
End If

If Check2.Value = 1 Then
 Randomize
    pid = Int(ListFile.ListCount * Rnd)
MediaPlayer1.Filename = ListFile.List(pid)
Exit Sub
End If

End Sub
Private Sub MediaPlayer1_NewStream()


If pid + 1 <= ListFile.ListCount Then
  CooLine1.Display = "[" + Str(pid + 1) + " -" + Str(ListFile.ListCount) + " ] " + MediaPlayer1.Filename & "  - Snowman Media  3.0"
      ListFile.ListIndex = pid
End If
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
Sub reco2()
Picture19.BackColor = &HFFFF&
Label29.BackColor = &HFFFF&
Label29.ForeColor = &HFF0000
Picture20.BackColor = &HFFFF&
Label30.BackColor = &HFFFF&
Label30.ForeColor = &HFF0000
Picture21.BackColor = &HFFFF&
Label31.BackColor = &HFFFF&
Label31.ForeColor = &HFF0000
Picture26.BackColor = &HFFFF&
Label43.BackColor = &HFFFF&
Label43.ForeColor = &HFF0000
Picture25.BackColor = &HFFFF&
Label42.BackColor = &HFFFF&
Label42.ForeColor = &HFF0000
Picture24.BackColor = &HFFFF&
Label41.BackColor = &HFFFF&
Label41.ForeColor = &HFF0000
Picture22.BackColor = &HFFFF&
Label40.BackColor = &HFFFF&
Label40.ForeColor = &HFF0000
End Sub

Private Sub Picture18_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Call reco2
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


Private Sub RA_OnClipClosed()
If RmStop = True Then
RA.Left = 10000
SLD1.Left = 10000
     
     If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(3)) = 1 Then
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
MediaPlayer1.Filename = MediaPlayer1.Filename
Exit Sub
End If

If Check1.Value = 0 And Check2.Value = 0 Then
pid = pid + 1
If pid = ListFile.ListCount Then
MediaPlayer1.Filename = "LLXX"
     Form4.Asd = True
End If
MediaPlayer1.Filename = ListFile.List(pid)
Exit Sub
End If

If Check1.Value = 1 Then
pid = pid + 1
If pid = ListFile.ListCount Then
pid = 0
End If
MediaPlayer1.Filename = ListFile.List(pid)
Exit Sub
End If

If Check2.Value = 1 Then
 Randomize
    pid = Int(ListFile.ListCount * Rnd)
MediaPlayer1.Filename = ListFile.List(pid)
Exit Sub
End If
End If
RmStop = True
End Sub

Private Sub RA_OnClipOpened(ByVal shortClipName As String, ByVal url As String)

If pid + 1 <= ListFile.ListCount Then
  CooLine1.Display = "[" + Str(pid + 1) + " -" + Str(ListFile.ListCount) + " ] " + MediaPlayer1.Filename & "  - Snowman Media  3.0"
     ListFile.ListIndex = pid
End If
SLD1.Max = RA.GetLength
Label1.Caption = "rm"
RA.Left = 0
SLD1.Left = -45
MediaPlayer1.Filename = "LLXX"
SendMCIString "stop cd wait", True
Command = "seek cd to " & Track
SendMCIString Command, True
Playing = False
Update
End Sub

Private Sub RA_OnPositionChange(ByVal lPos As Long, ByVal lLen As Long)
If RmGn = True Then SLD1.Value = RA.GetPosition
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

Private Sub TrackSelection_Click()
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
If RA.Left = 0 Then RmStop = False
If Len(MediaPlayer1.Filename) > 0 Then MediaPlayer1.Filename = "LLXX"
If RA.Left = 0 Then RmStop = False
RA.DoStop
RA.Left = 10000
SLD1.Left = 10000
End If
    
    
End Sub
Private Sub Volume_Change()
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
        Eject.Enabled = True
        Form102.Image11.ToolTipText = "弹出光驱  - Snowman Media  2.0"
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
     Eject.Enabled = False
     Form102.Image11.ToolTipText = "送入光驱  - Snowman Media  2.0"
     If (CDLoad = True) Then
        CDLoad = False
        Playing = False
        TrackTime.Caption = ""
        TrackTime.Caption = ""
        TimeWindow.Text = ""
    End If
End If
End Sub
Private Sub timer1_timer()
rgt = Time
  Update
AloT = AloT + 1
If AloT > 600 Then
'If AloT > 10 Then
  If Len(MediaPlayer1.Filename) > 0 And LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Op_" + Str(5)) = True Then Call SetAlo
AloT = 0
End If

If MediaPlayer1.DisplaySize <> mpFullScreen Then
If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(8)) = 1 Then SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If

'If Len(MediaPlayer1.Filename) > 0 Then Label1.Caption = "media"
'If Playing = True Then Label1.Caption = "cd"


End Sub

Sub Tiday()
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
          If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(8)) = 1 Then
          SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
          Else
             SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
          End If
          Image10.Picture = Image1.Picture
          SkinForm1.SkinPath = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Skin_Path")
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
          
    End Select
   
End Sub

 Private Sub ctListBar1_ItemClick(ByVal nList As Integer, ByVal nItem As Integer)
On Error Resume Next
 Dim Result As Long
Select Case nList
        Case 1
          Select Case nItem
          Case 1
                    If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Ch_" + Str(9)) = 1 Then
          SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
          Else
             SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
          End If

          SkinForm1.SkinPath = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Sflake_Path")
          MediaPlayer1.ShowControls = False
          MediaPlayer1.ShowStatusBar = False
          CooLine1.Left = myReadINI((LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Sflake_Path")) + "\sflake_info.sfl", "label", "x", "") '+ 25
          CooLine1.Top = myReadINI((LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Sflake_Path")) + "\sflake_info.sfl", "label", "y", "") '+ 25
          CooLine1.Width = myReadINI((LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Sflake_Path")) + "\sflake_info.sfl", "label", "w", "")
          Frame12.Left = 10000
          Frame4.Left = 10000
          ctListBar1.Left = 10000
          Frame2.Top = myReadINI((LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Sflake_Path")) + "\sflake_info.sfl", "video", "y", "") '+ 25
          Frame2.Left = myReadINI((LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Sflake_Path")) + "\sflake_info.sfl", "video", "x", "") '+ 25
          Frame2.Width = myReadINI((LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Sflake_Path")) + "\sflake_info.sfl", "video", "w", "")
          Frame2.Height = myReadINI((LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Sflake_Path")) + "\sflake_info.sfl", "video", "h", "")
          Frame5.Left = 10000
          Frame6.Left = 10000
          Frame7.Left = 10000
          MediaPlayer1.Top = 0
          MediaPlayer1.Top = 0
          MediaPlayer1.Width = Frame2.Width
          MediaPlayer1.Height = Frame2.Height
          Image10.Picture = LoadPicture(LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Sflake_Path") + "\sflake_logo.bmp")
          
          
                   Case 2
        If Len(MediaPlayer1.Filename) > 0 Then
        On Error Resume Next
        Form200.Show
        jn = MediaPlayer1.Filename
        jd = MediaPlayer1.CurrentPosition
        MediaPlayer1.Filename = "LLXX"
        Form200.MediaPlayer1.Filename = jn
       Form200.MediaPlayer1.CurrentPosition = jd
       
        f200 = 1
        Else: MsgBox ("无可用视频,请先选定视频或图片文件再打开视频窗口.")
        End If
        Case 3
        If Len(MediaPlayer1.Filename) > 0 Then
 SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
        MediaPlayer1.DisplaySize = mpFullScreen
          Else: MsgBox ("无可用视频,请先选定视频或图片文件再开始全屏欣赏.")
          End If
        Case 4
        If Len(MediaPlayer1.Filename) > 0 Then
        If Label26.Caption = "locked" Then
        Label26.Caption = "unlock"
        MsgBox ("已经解除对以下曲目的锁定: " & MediaPlayer1.Filename & " .")
        Exit Sub
        End If
        If Label26.Caption = "unlock" Then
        Label26.Caption = "locked"
        MsgBox ("已经锁定播放以下曲目: " & MediaPlayer1.Filename & " .")
        Exit Sub
        End If
        Else: MsgBox ("还没选择要锁定的曲目,请选定好曲目后再锁定.")
        End If
            Case 5
            If Len(MediaPlayer1.Filename) > 0 Then
            Form5.Show
            Else
            MsgBox ("当前没有正在播放的媒体,无法标记媒体书签.请在正在播放媒体文件时再作尝试.")
            End If
            Case 6
            If Len(MediaPlayer1.Filename) = 0 Then
                     MsgBox ("找不到可以显示媒体说明的媒体文件,请确认该媒体含有说明部分后重试.")
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
        'Dim Result As Long
          Dim fName As String
      If Len(MediaPlayer1.Filename) > 0 Then
     fName = MediaPlayer1.Filename
     Result = ShowProperties(fName, Me.hwnd)
        Else: MsgBox ("没有找到可以显示媒体属性的媒体文件,请先选择好媒体文件再要求显示媒体属性.")
        End If
        Case 8
        On Error Resume Next
        MediaPlayer1.ShowDialog mpShowDialogStatistics

        End Select
      
        
        Case 3
            Select Case nItem
                Case 1
          CommonDialog1.Filter = "媒体文件:Mp3、Wma、Wav、Wax、Asf、Rmi、Asx、Mov、M1v、Mp2、Mpg、Mpeg、Mpa、Mpe、Avi、Mid、Qt、M3u、Aif、Aifc、Aiff、Au、Snd、Bmp、Jpg、Did、Wmf、Gif、Rle、Cur、Emf..." & _
          "|*.au;*.and;*.aif;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.m3u;*.wma;*.wav;*.avi;*.mid;*.bmp;*.jpg;*.did;*.wmf;*.gif;*.rle;*.cur;*.emf|所有文件:*.*|*.*"
          CommonDialog1.FilterIndex = 1
          CommonDialog1.Filename = ""
          CommonDialog1.ShowOpen
          If Len(CommonDialog1.Filename) > 0 Then
         ListFile.AddItem CommonDialog1.Filename, ListFile.ListCount
         MediaPlayer1.Filename = ListFile.List(ListFile.ListCount - 1)
           Label1.Caption = "media"
           TrackSelection.Left = 10000
           Frame8.Left = 2880
           Frame1.Left = 10000
          End If
          Case 5
        Shell (Label3.Caption + "\SmM_FP.exe")
        Case 2
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
  If FileExists(Label27.Caption + "MPEGAV\") = True Then
          File1.Path = Label27.Caption + "MPEGAV\"
                 File1.Pattern = "*.au;*.dat;*.and;*.aif;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.m3u;*.wma;*.wav;*.avi;*.mid;*.bmp;*.jpg;*.did;*.wmf;*.gif;*.rle;*.cur;*.emf"
          Dim ind As Integer
              For ind = 0 To File1.ListCount - 1
             ListFile.AddItem Label27.Caption + "MPEGAV\" + File1.List(ind), ListFile.ListCount
                  Next ind
            pid = ListFile.ListCount - File1.ListCount
                    MediaPlayer1.Filename = ListFile.List(pid)
         CooLine1.Display = "[ VCD Video ] MediaPlayer1.Filename  - Snowman Media  3.0"
        Label1.Caption = "media"
        TrackSelection.Left = 10000
       Frame8.Left = 2880
        Frame1.Left = 10000
        Else: MsgBox ("光盘中没有VCD光盘,请先放入VCD光盘再使用本功能.")
        End If
       Case 3
  
           On Error Resume Next
           pIda = Shell(Label3.Caption + "\SmM_DP.exe", vbNormalFocus)
          pHnd = OpenProcess(SYNCHRONIZE, 0, pIda)
          If pHnd <> 0 Then
            Call WaitForSingleObject(pHnd, INFINITE)
           Call CloseHandle(pHnd)
           End If
       Case 4
      Dim firstpath As String, dircount As Integer, NumFiles As Integer
    Cdno = 0
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
               Dir1.Path = Label27.Caption
               File1.Pattern = "*.au;*.and;*.aif;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.m3u;*.wma;*.wav;*.avi;*.mid"
               If File1.Path = Label3.Caption Then GoTo cc
                 firstpath = Dir1.Path
                dircount = Dir1.ListCount
                 NumFiles = 0                       ' Reset global foundfiles indicator.
               Result = DirDiver(firstpath, dircount, "")
                  File1.Path = Dir1.Path
                       If Cdno = 0 Then
cc:
                     MsgBox ("在光驱中找不到含有媒体文件的光盘.请放入含有媒体文件的光盘后重试.")
                     Exit Sub
                     End If
                    pid = ListFile.ListCount - Cdno
              MediaPlayer1.Filename = ListFile.List(pid)
            Cdno = 0
                Label1.Caption = "media"
        TrackSelection.Left = 10000
       Frame8.Left = 2880
        Frame1.Left = 10000
   
   
   
   
   
   Case 6
          Shell (Label3.Caption + "\SmM_3P.exe")
       Case 7
        pid = -1
       MediaPlayer1.Filename = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Name")
       MediaPlayer1.CurrentPosition = LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Rute")
       Dim idc As Integer
            For idc = 0 To ListFile.ListCount - 1
            If ListFile.List(idc) = MediaPlayer1.Filename Then
            pid = idc
            ListFile.ListIndex = pid
            Exit Sub
            End If
            Next
          If pid = -1 Then
          ListFile.AddItem (MediaPlayer1.Filename)
          pid = ListFile.ListCount - 1
          ListFile.ListIndex = pid
          End If
        Case 8
        If LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Name_A") <> "Error" Then
          pid = -1
       MediaPlayer1.Filename = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Name_A")
       MediaPlayer1.CurrentPosition = LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Rute_A")
            For idc = 0 To ListFile.ListCount
            If ListFile.List(idc) = MediaPlayer1.Filename Then
            pid = idc
            ListFile.ListIndex = pid
            Exit Sub
            End If
            Next
          If pid = -1 Then
          ListFile.AddItem (MediaPlayer1.Filename)
          pid = ListFile.ListCount - 1
          ListFile.ListIndex = pid
          End If
         Else: MsgBox ("本媒体书签为空,没有可以用于播放的媒体文件记录.请在标记本书签后再次尝试.")
         End If
       Case 9
           If LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Name_B") <> "Error" Then
          pid = -1
       MediaPlayer1.Filename = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Name_B")
       MediaPlayer1.CurrentPosition = LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Rute_B")
       'Dim idc As Integer
            For idc = 0 To ListFile.ListCount
            If ListFile.List(idc) = MediaPlayer1.Filename Then
            pid = idc
            ListFile.ListIndex = pid
            Exit Sub
            End If
            Next
          If pid = -1 Then
          ListFile.AddItem (MediaPlayer1.Filename)
          pid = ListFile.ListCount - 1
          ListFile.ListIndex = pid
          End If
         Else: MsgBox ("本媒体书签为空,没有可以用于播放的媒体文件记录.请在标记本书签后再次尝试.")
         End If
        Case 10
        If LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Name_C") <> "Error" Then
          pid = -1
       MediaPlayer1.Filename = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Name_C")
       MediaPlayer1.CurrentPosition = LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Rute_C")
       'Dim idc As Integer
            For idc = 0 To ListFile.ListCount
            If ListFile.List(idc) = MediaPlayer1.Filename Then
            pid = idc
            ListFile.ListIndex = pid
            Exit Sub
            End If
            Next
          If pid = -1 Then
          ListFile.AddItem (MediaPlayer1.Filename)
          pid = ListFile.ListCount - 1
          ListFile.ListIndex = pid
          End If
         Else: MsgBox ("本媒体书签为空,没有可以用于播放的媒体文件记录.请在标记本书签后再次尝试.")
         End If
       Case 11
          If LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Name_D") <> "Error" Then
          pid = -1
       MediaPlayer1.Filename = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Name_D")
       MediaPlayer1.CurrentPosition = LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Alo_Rute_D")
       'Dim idc As Integer
            For idc = 0 To ListFile.ListCount
            If ListFile.List(idc) = MediaPlayer1.Filename Then
            pid = idc
            ListFile.ListIndex = pid
            Exit Sub
            End If
            Next
          If pid = -1 Then
          ListFile.AddItem (MediaPlayer1.Filename)
          pid = ListFile.ListCount - 1
          ListFile.ListIndex = pid
          End If
         Else: MsgBox ("本媒体书签为空,没有可以用于播放的媒体文件记录.请在标记本书签后再次尝试.")
         End If
         End Select
        
         
         
    
      Case 2
          Select Case nItem
             Case 1
           CommonDialog1.Filter = "媒体文件:Mp3、Wma、Wav、Wax、Asf、Rmi、Asx、Mov、M1v、Mp2、Mpg、Mpeg、Mpa、Mpe、Avi、Mid、Qt、M3u、Aif、Aifc、Aiff、Au、Snd、Bmp、Jpg、Did、Wmf、Gif、Rle、Cur、Emf..." & _
          "|*.au;*.and;*.aif;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.m3u;*.wma;*.wav;*.avi;*.mid;*.bmp;*.jpg;*.did;*.wmf;*.gif;*.rle;*.cur;*.emf|所有文件:*.*|*.*"
          CommonDialog1.FilterIndex = 1
          CommonDialog1.Filename = ""
          CommonDialog1.ShowOpen
          If Len(CommonDialog1.Filename) > 0 Then
       SelectFileName = CommonDialog1.Filename
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
       AddUrl = InputBox("请输入你要加入到列表的媒体文件的 URL 地址.地址可以是 Internet 上的也可以是本地主机的,Snowman Media  3.0 将自动识别并进行播放.")
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
                CommonDialog1.Filename = ""
               CommonDialog1.Filter = "列表文件:M3u" & _
          "|*.M3u|所有文件:*.*|*.*"
             CommonDialog1.ShowOpen
             If Len(CommonDialog1.Filename) > 0 Then
              Open CommonDialog1.Filename For Input As #1
           While Not EOF(1)
          Line Input #1, test
           ListFile.AddItem RTrim(test)
           Wend
             Close #1
            End If
            Case 4
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
    .lpszTitle = "请选择你要添加媒体文件夹.当你确定后Snomwan Media 3.0将为你打开它.文件夹内所有文件将自动加入列表."
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
               
                 Dir1.Path = txtPath
               File1.Pattern = "*.au;*.and;*.aif;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.m3u;*.wma;*.wav;*.avi;*.mid"
                 firstpath = Dir1.Path
                dircount = Dir1.ListCount
                 NumFiles = 0                       ' Reset global foundfiles indicator.
               Result = DirDiver(firstpath, dircount, "")
                  File1.Path = Dir1.Path
    Case 8
                ListFile.Clear
                  Case 7
            If ListFile.ListCount > 0 Then
                  Dim i As Integer
              CommonDialog1.Filename = ""
               CommonDialog1.Filter = "列表文件:M3u" & _
          "|*.m3u|所有文件:*.*|*.*"
             CommonDialog1.ShowSave
             If Len(CommonDialog1.Filename) > 0 Then
             Open CommonDialog1.Filename For Output As #1
    For i = 0 To ListFile.ListCount - 1
     Print #1, ListFile.List(i)
    Next i
   Close (1)
   End If
        Else: MsgBox ("当前播放列表为空无法保存,请先加入曲目到列表.")
        End If
      Case 6
      Call Tiday
      
      
    End Select
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      
      Case 4
      Select Case nItem
        
                         Case 1
                 If LyfTools1.IsConnected Then
                  MsgBox ("访问 流动网络H2ont媒体指南 时需要连接网络并从相关站点下载媒体资料.播放器检测到你的计算机当前并没有连接网络,请在确保已经连接上网络后重试.")
               End If
        Case 2
         If LyfTools1.IsConnected Then
             Label2.Caption = "media"
             Formo.Show
             Else
               Label2.Caption = "media"
            Form2.Show
            End If
          Case 3
           If LyfTools1.IsConnected Then
          Label2.Caption = "flash"
             Formo.Show
                  Else
                      Label2.Caption = "flash"
                  Form2.Show
            End If
            End Select
     Case 5
        Select Case nItem

                 Case 1
               i = 0
                 CommonDialog1.Filename = ""
               CommonDialog1.Filter = "列表文件:M3u" & _
          "|*.M3u|所有文件:*.*|*.*"
             CommonDialog1.ShowOpen
             If Len(CommonDialog1.Filename) > 0 Then
              Open CommonDialog1.Filename For Input As #1
                
           While Not EOF(1)
          Line Input #1, test
           ListFile.AddItem RTrim(test)
            i = i + 1
           Wend
             Close #1
             pid = ListFile.ListCount - i
             MediaPlayer1.Filename = ListFile.List(pid)
                End If
                   Case 2
            Call ToolB
                        Case 4
             Call ToolC
                If FileExists(Label3.Caption + "\SmMDb.dat") = False Then
                 MsgBox ("你是第一次访问媒体收藏中的所有媒体.使用本功能时需要调用计算机中的媒体资源.为了搜集这些资源以便访问,要求更新媒体库.当你按下确定后将自动更新.过程需要一段时间,你可以选择停止等有时间再更新.")
                 Call Find
               Else
                Open Label3.Caption + "\SmMDb.dat" For Input As #1
                           

                 
           While Not EOF(1)
          Line Input #1, test
           List2.AddItem RTrim(test)
           Wend
             Close #1
             Label39.Caption = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "MediaWH_PS")
               Label32.Caption = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "MediaWH_TM")
                Label36.Caption = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "MediaWH_DA")
                 Label46.Caption = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "MediaWH_DY")

                    Label37.Caption = Str(List2.ListCount)
                End If
          Case 3
             With BI
              .hOwner = Me.hwnd
            nFolder = GetFolderValue(m_wCurOptIdx)
            If SHGetSpecialFolderLocation(ByVal Me.hwnd, ByVal nFolder, IDL) = noerror Then
            .pidlRoot = IDL.mkid.cb
             End If
     .pszDisplayName = String$(MAX_PATH, 0)
    .lpszTitle = "请选择你要播放的媒体文件夹.当你确定后 Snomwan Media 3.0 将为你打开它,并自动播放其中的所有媒体."
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
                             'File1.Path = txtPath
                  Dir1.Path = txtPath
               File1.Pattern = "*.au;*.and;*.aif;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.m3u;*.wma;*.wav;*.avi;*.mid"
                 firstpath = Dir1.Path
                dircount = Dir1.ListCount
                 NumFiles = 0                       ' Reset global foundfiles indicator.
               Result = DirDiver(firstpath, dircount, "")
                  File1.Path = Dir1.Path
                    pid = ListFile.ListCount - Cdno
              MediaPlayer1.Filename = ListFile.List(pid)
                  
                  
                  
                  
                  
                  
                  
                
           
           
           
           
           
          
        End Select
  Case 6
        Select Case nItem
               Case 1
           Shell (Label3.Caption + "\SmM_DE.exe")
               Case 2
           Shell (Label3.Caption + "\SmM_PB.exe")
               Case 3
           Form100.Show
               Case 4
           Form4.Show
                Case 5
           Shell (Label3.Caption + "\SmM_St.exe")
         End Select
  Case 7
       Select Case nItem
        Case 1
           Shell (Label3.Caption + "\SmM_Hp.exe")
        Case 2
           Agent1.Characters.Load "merlin.acs", DATAPATH
           Set merlin = Agent1.Characters("merlin.acs")
           merlin.Show
           'merlin.Think ("Bus")
           'i = merlin.GestureAt(100, 500)
           merlin.MoveTo Me.Left / 15, Me.Top / 15
           merlin.Speak "欢迎使用 Snowman Media  3.0!尽情享受她为你带来的愉悦多媒体体验吧!"
        End Select
End Select
End Sub
