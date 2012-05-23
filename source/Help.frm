VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "LYFTOOLS.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sm.M. Help"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   6180
   StartUpPosition =   2  '屏幕中心
   Begin API控制大全.LyfTools LyfTools1 
      Left            =   10935
      Top             =   4050
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2565
      ScaleHeight     =   255
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   1170
      Width           =   1140
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "便捷操作"
         ForeColor       =   &H00FF0000&
         Height          =   465
         Left            =   180
         TabIndex        =   3
         Top             =   45
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1305
      ScaleHeight     =   255
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   1170
      Width           =   1140
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "快速入门"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   180
         TabIndex        =   1
         Top             =   45
         Width           =   780
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -135
      TabIndex        =   4
      Top             =   1125
      Width           =   6810
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   180
         ScaleHeight     =   255
         ScaleWidth      =   1110
         TabIndex        =   9
         Top             =   45
         Width           =   1140
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "了解 Sm.M."
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   90
            TabIndex        =   10
            Top             =   45
            Width           =   1185
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5220
         ScaleHeight     =   255
         ScaleWidth      =   1020
         TabIndex        =   7
         Top             =   45
         Width           =   1050
         Begin VB.Label Label19 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "我要提问"
            ForeColor       =   &H00FF0000&
            Height          =   465
            Left            =   135
            TabIndex        =   8
            Top             =   45
            Width           =   1365
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3960
         ScaleHeight     =   255
         ScaleWidth      =   1110
         TabIndex        =   5
         Top             =   45
         Width           =   1140
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "了解媒体"
            ForeColor       =   &H00FF0000&
            Height          =   465
            Left            =   180
            TabIndex        =   6
            Top             =   45
            Width           =   1365
         End
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6945
      Left            =   -45
      TabIndex        =   11
      Top             =   1215
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   12250
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      BackColor       =   0
      TabCaption(0)   =   "A"
      TabPicture(0)   =   "Help.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "B"
      TabPicture(1)   =   "Help.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "C"
      TabPicture(2)   =   "Help.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "D"
      TabPicture(3)   =   "Help.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "E"
      TabPicture(4)   =   "Help.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame2"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "F"
      TabPicture(5)   =   "Help.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame7"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Tab 6"
      TabPicture(6)   =   "Help.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame8"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Tab 7"
      TabPicture(7)   =   "Help.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame9"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).ControlCount=   1
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   -74955
         TabIndex        =   69
         Top             =   180
         Width           =   6900
         Begin VB.Label Label18 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "leask@21cn.com"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   4050
            MouseIcon       =   "Help.frx":03EA
            MousePointer    =   99  'Custom
            TabIndex        =   71
            Top             =   2025
            Width           =   1260
         End
         Begin VB.Label Label24 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   $"Help.frx":053C
            ForeColor       =   &H00000000&
            Height          =   735
            Left            =   495
            TabIndex        =   70
            Top             =   1170
            Width           =   5010
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   -74955
         TabIndex        =   66
         Top             =   225
         Width           =   6900
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   3660
            Left            =   270
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   67
            Text            =   "Help.frx":05D4
            Top             =   585
            Width           =   5595
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "流动网络"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   1530
            MouseIcon       =   "Help.frx":0A09
            MousePointer    =   99  'Custom
            TabIndex        =   73
            Top             =   4815
            Width           =   720
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   $"Help.frx":0B5B
            ForeColor       =   &H000000FF&
            Height          =   870
            Left            =   270
            TabIndex        =   72
            Top             =   4455
            Width           =   5595
         End
         Begin VB.Label Label22 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "了解 Snowman Media 支持的媒体格式"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   270
            TabIndex        =   68
            Top             =   225
            Width           =   5280
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   -74955
         TabIndex        =   62
         Top             =   225
         Width           =   6900
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   4065
            Left            =   315
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   65
            Text            =   "Help.frx":0C15
            Top             =   720
            Width           =   5550
         End
         Begin VB.Label Label21 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "为了有方便地使用 Snowman Media 在设计时预设了下列快捷键"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   270
            TabIndex        =   64
            Top             =   225
            Width           =   5280
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "返回"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   5175
            TabIndex        =   63
            Top             =   3645
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   -74955
         TabIndex        =   57
         Top             =   270
         Width           =   6900
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "返回"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   5130
            MouseIcon       =   "Help.frx":129C
            MousePointer    =   99  'Custom
            TabIndex        =   60
            Top             =   3645
            Width           =   1365
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "使用同步媒体列表方便播放媒体"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   270
            TabIndex        =   59
            Top             =   180
            Width           =   4740
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   $"Help.frx":13EE
            ForeColor       =   &H80000008&
            Height          =   2400
            Left            =   405
            TabIndex        =   58
            Top             =   810
            Width           =   5370
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   6225
         Left            =   -75000
         TabIndex        =   31
         Top             =   270
         Width           =   6900
         Begin VB.Label Label54 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "系统要求"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   5355
            MouseIcon       =   "Help.frx":16BE
            MousePointer    =   99  'Custom
            TabIndex        =   49
            Top             =   270
            Width           =   1050
         End
         Begin VB.Line Line27 
            X1              =   2475
            X2              =   2475
            Y1              =   2880
            Y2              =   3915
         End
         Begin VB.Line Line26 
            X1              =   2205
            X2              =   2475
            Y1              =   3915
            Y2              =   3915
         End
         Begin VB.Line Line25 
            X1              =   2205
            X2              =   2205
            Y1              =   3915
            Y2              =   4680
         End
         Begin VB.Shape Shape10 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00FF0000&
            Height          =   870
            Left            =   2385
            Top             =   4140
            Width           =   1230
         End
         Begin VB.Line Line24 
            X1              =   3645
            X2              =   3645
            Y1              =   3870
            Y2              =   4860
         End
         Begin VB.Line Line23 
            X1              =   3195
            X2              =   3645
            Y1              =   3870
            Y2              =   3870
         End
         Begin VB.Line Line22 
            X1              =   3330
            X2              =   3780
            Y1              =   3780
            Y2              =   3780
         End
         Begin VB.Line Line21 
            X1              =   3330
            X2              =   3330
            Y1              =   2880
            Y2              =   3780
         End
         Begin VB.Line Line20 
            X1              =   2565
            X2              =   3330
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line19 
            X1              =   2565
            X2              =   2565
            Y1              =   2700
            Y2              =   2880
         End
         Begin VB.Line Line18 
            X1              =   3780
            X2              =   3645
            Y1              =   4860
            Y2              =   4860
         End
         Begin VB.Shape Shape9 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00FF0000&
            Height          =   285
            Left            =   3825
            Top             =   4725
            Width           =   2265
         End
         Begin VB.Label Label52 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "音量控制 "
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   3870
            MouseIcon       =   "Help.frx":1810
            TabIndex        =   48
            Top             =   4770
            Width           =   750
         End
         Begin VB.Shape Shape8 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00FF0000&
            Height          =   1050
            Left            =   3825
            Top             =   3645
            Width           =   2265
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "状态栏"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   2430
            MouseIcon       =   "Help.frx":1962
            TabIndex        =   47
            Top             =   4185
            Width           =   540
         End
         Begin VB.Line Line17 
            X1              =   2205
            X2              =   2340
            Y1              =   4680
            Y2              =   4680
         End
         Begin VB.Line Line16 
            X1              =   3180
            X2              =   3180
            Y1              =   3330
            Y2              =   3870
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "详细介绍"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   1035
            MouseIcon       =   "Help.frx":1AB4
            MousePointer    =   99  'Custom
            TabIndex        =   46
            Top             =   4050
            Width           =   720
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00FF0000&
            Height          =   690
            Left            =   360
            Top             =   3600
            Width           =   1725
         End
         Begin VB.Label Label50 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "同步媒体列表"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   945
            MouseIcon       =   "Help.frx":1C06
            TabIndex        =   45
            Top             =   3645
            Width           =   1080
         End
         Begin VB.Line Line15 
            X1              =   2115
            X2              =   2295
            Y1              =   3735
            Y2              =   3735
         End
         Begin VB.Line Line14 
            X1              =   2295
            X2              =   2295
            Y1              =   3735
            Y2              =   3240
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00FF0000&
            Height          =   690
            Left            =   450
            Top             =   4320
            Width           =   1725
         End
         Begin VB.Shape Shape5 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00FF0000&
            Height          =   690
            Left            =   3825
            Top             =   2925
            Width           =   2265
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00FF0000&
            Height          =   690
            Left            =   3825
            Top             =   1710
            Width           =   2265
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00FF0000&
            Height          =   510
            Left            =   3825
            Top             =   1170
            Width           =   2265
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00FF0000&
            Height          =   645
            Left            =   3825
            Top             =   495
            Width           =   2265
         End
         Begin VB.Line Line13 
            X1              =   270
            X2              =   270
            Y1              =   3015
            Y2              =   4410
         End
         Begin VB.Label Label46 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "模式选择按钮"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   3870
            MouseIcon       =   "Help.frx":1D58
            TabIndex        =   42
            Top             =   2970
            Width           =   1080
         End
         Begin VB.Line Line12 
            X1              =   3465
            X2              =   3780
            Y1              =   3060
            Y2              =   3060
         End
         Begin VB.Line Line11 
            X1              =   3465
            X2              =   3465
            Y1              =   2700
            Y2              =   3060
         End
         Begin VB.Line Line10 
            X1              =   3015
            X2              =   3465
            Y1              =   2700
            Y2              =   2700
         End
         Begin VB.Label Label44 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "功能菜单"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   495
            MouseIcon       =   "Help.frx":1EAA
            TabIndex        =   41
            Top             =   4365
            Width           =   720
         End
         Begin VB.Line Line9 
            X1              =   270
            X2              =   405
            Y1              =   4410
            Y2              =   4410
         End
         Begin VB.Label Label40 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "播放控制"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   3870
            MouseIcon       =   "Help.frx":1FFC
            TabIndex        =   40
            Top             =   2475
            Width           =   720
         End
         Begin VB.Label Label39 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "用于对当前媒体作跳跃式进度搜索"
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   3870
            TabIndex        =   39
            Top             =   1980
            Width           =   2205
         End
         Begin VB.Label Label37 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "搜索栏"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   3870
            MouseIcon       =   "Help.frx":214E
            TabIndex        =   38
            Top             =   1755
            Width           =   540
         End
         Begin VB.Line Line8 
            X1              =   3465
            X2              =   3735
            Y1              =   1845
            Y2              =   1845
         End
         Begin VB.Line Line7 
            X1              =   3465
            X2              =   3465
            Y1              =   2520
            Y2              =   1845
         End
         Begin VB.Line Line6 
            X1              =   2745
            X2              =   3465
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Line Line5 
            X1              =   1755
            X2              =   3780
            Y1              =   2565
            Y2              =   2565
         End
         Begin VB.Line Line4 
            X1              =   1755
            X2              =   1755
            Y1              =   2700
            Y2              =   2565
         End
         Begin VB.Label Label34 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "用于播放显示视频媒体"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   3870
            TabIndex        =   37
            Top             =   1440
            Width           =   2205
         End
         Begin VB.Label Label31 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "嵌入式视频窗口"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   3870
            MouseIcon       =   "Help.frx":22A0
            TabIndex        =   36
            Top             =   1215
            Width           =   1260
         End
         Begin VB.Line Line3 
            X1              =   2970
            X2              =   3780
            Y1              =   1305
            Y2              =   1305
         End
         Begin VB.Line Line2 
            X1              =   2970
            X2              =   3780
            Y1              =   630
            Y2              =   630
         End
         Begin VB.Label Label33 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "这个栏目教你如何使用 Snowman Media 播放媒体"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   225
            TabIndex        =   35
            Top             =   90
            Width           =   3870
         End
         Begin VB.Label Label26 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "信息栏"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   3870
            MouseIcon       =   "Help.frx":23F2
            TabIndex        =   34
            Top             =   540
            Width           =   540
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "详细介绍"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   1125
            MouseIcon       =   "Help.frx":2544
            MousePointer    =   99  'Custom
            TabIndex        =   32
            Top             =   4770
            Width           =   720
         End
         Begin VB.Label Label30 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "用于显示正在播放的媒体的详细信息"
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   3870
            TabIndex        =   51
            Top             =   765
            Width           =   2205
         End
         Begin VB.Label Label47 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "用于选择当前是播放媒体文件还是 CD 音频"
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   3870
            TabIndex        =   43
            Top             =   3195
            Width           =   2205
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00C00000&
            Height          =   465
            Left            =   3825
            Top             =   2430
            Width           =   2265
         End
         Begin VB.Label Label41 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "用于控制媒体播放"
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   3870
            TabIndex        =   50
            Top             =   2700
            Width           =   2205
         End
         Begin VB.Label Label48 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "循环方式/CD 曲目点播"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   3870
            MouseIcon       =   "Help.frx":2696
            TabIndex        =   44
            Top             =   3690
            Width           =   1800
         End
         Begin VB.Label Label53 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "用于控制音量"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   4770
            TabIndex        =   55
            Top             =   4770
            Width           =   1665
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "用于显示当前媒体播放状态和播放时间"
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   2430
            TabIndex        =   53
            Top             =   4410
            Width           =   1170
         End
         Begin VB.Label Label49 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "用于打开调用预设功能,请看详细介绍"
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   495
            TabIndex        =   33
            Top             =   4590
            Width           =   1665
         End
         Begin VB.Label Label51 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "用于管理媒体播放任务,请看详细介绍"
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   405
            TabIndex        =   54
            Top             =   3870
            Width           =   1620
         End
         Begin VB.Image Image1 
            Height          =   3135
            Left            =   180
            Picture         =   "Help.frx":27E8
            Top             =   405
            Width           =   3090
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "开始前请先了解系统要求"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   4095
            MouseIcon       =   "Help.frx":631C
            MousePointer    =   99  'Custom
            TabIndex        =   52
            Top             =   270
            Width           =   1980
         End
         Begin VB.Label Label45 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "在媒体模式时显示模式控制控制媒体播放的循环方式;在 CD 模式时显示 CD 曲目下拉选框"
            ForeColor       =   &H80000008&
            Height          =   900
            Left            =   3870
            TabIndex        =   56
            Top             =   3915
            Width           =   2205
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   -74955
         TabIndex        =   27
         Top             =   270
         Width           =   6900
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   2940
            Left            =   270
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   61
            Text            =   "Help.frx":646E
            Top             =   1485
            Width           =   5685
         End
         Begin VB.Label Label42 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   $"Help.frx":7027
            ForeColor       =   &H80000008&
            Height          =   1050
            Left            =   315
            TabIndex        =   30
            Top             =   315
            Width           =   5370
         End
         Begin VB.Label Label43 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "如何使用功能菜单"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   270
            TabIndex        =   29
            Top             =   90
            Width           =   4740
         End
         Begin VB.Label Label55 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "返回"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   5265
            MouseIcon       =   "Help.frx":712C
            MousePointer    =   99  'Custom
            TabIndex        =   28
            Top             =   4725
            Width           =   1365
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   -74955
         TabIndex        =   18
         Top             =   225
         Width           =   6900
         Begin VB.Label Label36 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   $"Help.frx":727E
            ForeColor       =   &H80000008&
            Height          =   1005
            Left            =   1620
            TabIndex        =   26
            Top             =   3780
            Width           =   4335
         End
         Begin VB.Label Label35 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "有关说明:"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   495
            TabIndex        =   25
            Top             =   3735
            Width           =   810
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "软件环境:"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   495
            TabIndex        =   24
            Top             =   2655
            Width           =   810
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   $"Help.frx":7372
            ForeColor       =   &H80000008&
            Height          =   960
            Left            =   1620
            TabIndex        =   23
            Top             =   2655
            Width           =   4470
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "硬件环境:"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   495
            TabIndex        =   22
            Top             =   495
            Width           =   810
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   $"Help.frx":7448
            ForeColor       =   &H80000008&
            Height          =   2130
            Left            =   1665
            TabIndex        =   21
            Top             =   495
            Width           =   3840
         End
         Begin VB.Label Label38 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "运行 Snowman Media 最少需要以下系统配置"
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   270
            TabIndex        =   20
            Top             =   135
            Width           =   4740
         End
         Begin VB.Label Label56 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "返回"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   5490
            MouseIcon       =   "Help.frx":750A
            MousePointer    =   99  'Custom
            TabIndex        =   19
            Top             =   4815
            Width           =   1365
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         ForeColor       =   &H00FFFFFF&
         Height          =   5460
         Left            =   0
         TabIndex        =   12
         Top             =   270
         Width           =   7080
         Begin VB.Line Line1 
            X1              =   450
            X2              =   5850
            Y1              =   4725
            Y2              =   4725
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright (C) 2000-2001 H2ont Leask"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   1575
            TabIndex        =   17
            Top             =   4860
            Width           =   3150
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "版本 HLS-SmM ilxz 3.5 Plus Edition"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   1485
            TabIndex        =   16
            Top             =   720
            Width           =   3060
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Snowman Media 多功能播放媒体及管理软件"
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   1485
            TabIndex        =   15
            Top             =   450
            Width           =   3420
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "通过本栏目对 Snowman Media 作初步了解."
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   315
            TabIndex        =   14
            Top             =   135
            Width           =   3975
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   $"Help.frx":765C
            ForeColor       =   &H00000000&
            Height          =   3210
            Left            =   270
            TabIndex        =   13
            Top             =   1125
            Width           =   5820
         End
      End
   End
   Begin VB.Image Image4 
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   4245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Image4.Picture = LoadPicture(App.Path + "\SmM_PT\SmM_ALA.gif")
Label1.Caption = "版本:  HLS-SmM " + LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "VolID")
End Sub

Private Sub Frame5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReCo
End Sub


Private Sub Label10_Click()
SSTab1.Tab = 1
End Sub

Private Sub Label11_Click()
SSTab1.Tab = 4
End Sub

Private Sub Label15_Click()
SSTab1.Tab = 1
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.BackColor = &HFF0000
Label15.ForeColor = &HFFFF&
End Sub

Private Sub Label16_Click()
SSTab1.Tab = 5
End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.BackColor = &HFF0000
Label16.ForeColor = &HFFFF&
End Sub

Private Sub Label17_Click()
SSTab1.Tab = 6
End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture4.BackColor = &HFF0000
Label17.ForeColor = &HFFFF&
End Sub

Private Sub Label18_Click()
        Me.LyfTools1.SendMail ("leask@21cn.com")
End Sub

Private Sub Label19_Click()
SSTab1.Tab = 7
End Sub

Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture6.BackColor = &HFF0000
Label19.ForeColor = &HFFFF&
End Sub

Private Sub Label20_Click()
SSTab1.Tab = 0
End Sub

Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture7.BackColor = &HFF0000
Label20.ForeColor = &HFFFF&
End Sub

Private Sub Label25_Click()
If Me.LyfTools1.IsConnected = True Then
        Me.LyfTools1.HttpTo ("http://www.h2ont.com")
Else
   MsgBox ("你还没有连接 Internet 无法打开网页,请连接网络后重试.")
End If

End Sub

Private Sub Label28_Click()
SSTab1.Tab = 3
End Sub

Private Sub Label54_Click()
SSTab1.Tab = 2
End Sub

Private Sub Label55_Click()
SSTab1.Tab = 1
End Sub

Private Sub Label56_Click()
SSTab1.Tab = 1
End Sub
Sub ReCo()
Picture2.BackColor = &HFFFF&
Label15.ForeColor = &HFF0000
Picture3.BackColor = &HFFFF&
Label16.ForeColor = &HFF0000
Picture4.BackColor = &HFFFF&
Label17.ForeColor = &HFF0000
Picture6.BackColor = &HFFFF&
Label19.ForeColor = &HFF0000
Picture7.BackColor = &HFFFF&
Label20.ForeColor = &HFF0000

End Sub
