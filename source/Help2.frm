VERSION 5.00
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "SmM_Tools.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snowman Media ilxz 帮助"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9885
   Icon            =   "Help2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9885
   StartUpPosition =   2  '屏幕中心
   Begin API控制大全.LyfTools Ly 
      Left            =   2295
      Top             =   7515
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   5910
      Left            =   0
      TabIndex        =   37
      Top             =   540
      Width           =   9870
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "系统要求"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   6705
         MouseIcon       =   "Help2.frx":2CFA
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   4635
         Width           =   735
      End
      Begin VB.Label Label37 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "在开始前你必须先了解你的计算机是否满足 Snowman Media ilxz 的系统要求。"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   1305
         TabIndex        =   51
         Top             =   4635
         Width           =   6300
      End
      Begin VB.Label Label42 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   $"Help2.frx":2E4C
         ForeColor       =   &H00000000&
         Height          =   3120
         Left            =   1035
         TabIndex        =   42
         Top             =   1395
         Width           =   8160
      End
      Begin VB.Label Label41 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "版本  ilxz 4.06.9142002"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   3375
         TabIndex        =   41
         Top             =   855
         Width           =   2070
      End
      Begin VB.Label Label40 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "通过本栏目对 Snowman Media ilxz 作初步了解。"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   180
         TabIndex        =   40
         Top             =   135
         Width           =   3975
      End
      Begin VB.Line Line37 
         X1              =   2160
         X2              =   7695
         Y1              =   5310
         Y2              =   5310
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (C) 2000-2002 H2O Networks"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   2160
         TabIndex        =   39
         Top             =   5445
         Width           =   5580
      End
      Begin VB.Label Label38 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Snowman Media ilxz"
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   2835
         TabIndex        =   38
         Top             =   540
         Width           =   1620
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   5880
      Left            =   0
      TabIndex        =   5
      Top             =   540
      Width           =   9870
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "    阅读本栏目可以快速掌握 Snowman Media ilxz 的基本使用方法。"
         ForeColor       =   &H00C00000&
         Height          =   630
         Left            =   180
         TabIndex        =   33
         Top             =   135
         Width           =   2340
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "音量控制"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   4905
         MouseIcon       =   "Help2.frx":3144
         TabIndex        =   30
         Top             =   5580
         Width           =   720
      End
      Begin VB.Line Line35 
         X1              =   5940
         X2              =   5850
         Y1              =   5625
         Y2              =   5625
      End
      Begin VB.Line Line34 
         X1              =   5940
         X2              =   5940
         Y1              =   4140
         Y2              =   5625
      End
      Begin VB.Line Line33 
         X1              =   4545
         X2              =   5940
         Y1              =   4140
         Y2              =   4140
      End
      Begin VB.Line Line30 
         X1              =   4545
         X2              =   4545
         Y1              =   3735
         Y2              =   4140
      End
      Begin VB.Shape Shape18 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FF0000&
         Height          =   330
         Left            =   3015
         Top             =   5490
         Width           =   2760
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "播放控制"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3510
         MouseIcon       =   "Help2.frx":3296
         TabIndex        =   28
         Top             =   4320
         Width           =   720
      End
      Begin VB.Line Line28 
         X1              =   3195
         X2              =   3285
         Y1              =   4365
         Y2              =   4365
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "集中管理播放机的各种功能和播放控制;共有九个功能分类。"
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   6705
         TabIndex        =   26
         Top             =   4905
         Width           =   2895
      End
      Begin VB.Label Label50 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Snowman Media 主窗口"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   630
         MouseIcon       =   "Help2.frx":33E8
         TabIndex        =   19
         Top             =   1125
         Width           =   1800
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FF0000&
         Height          =   1140
         Left            =   135
         Top             =   2475
         Width           =   2400
      End
      Begin VB.Line Line2 
         X1              =   2610
         X2              =   2655
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Line Line3 
         X1              =   2655
         X2              =   2655
         Y1              =   1170
         Y2              =   1665
      End
      Begin VB.Line Line4 
         X1              =   2610
         X2              =   3105
         Y1              =   2655
         Y2              =   2655
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FF0000&
         Height          =   1320
         Left            =   135
         Top             =   1035
         Width           =   2400
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "视频窗口"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   1710
         MouseIcon       =   "Help2.frx":353A
         TabIndex        =   18
         Top             =   2565
         Width           =   720
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FF0000&
         Height          =   1140
         Left            =   135
         Top             =   3735
         Width           =   2400
      End
      Begin VB.Line Line5 
         X1              =   2700
         X2              =   3195
         Y1              =   3510
         Y2              =   3510
      End
      Begin VB.Line Line6 
         X1              =   2700
         X2              =   2700
         Y1              =   3510
         Y2              =   3915
      End
      Begin VB.Line Line7 
         X1              =   2610
         X2              =   2700
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "进度栏"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   1890
         MouseIcon       =   "Help2.frx":368C
         TabIndex        =   16
         Top             =   3825
         Width           =   540
      End
      Begin VB.Line Line8 
         X1              =   3060
         X2              =   3060
         Y1              =   3870
         Y2              =   5130
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FF0000&
         Height          =   825
         Left            =   135
         Top             =   4995
         Width           =   2760
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "时间栏"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   2250
         MouseIcon       =   "Help2.frx":37DE
         TabIndex        =   14
         Top             =   5085
         Width           =   540
      End
      Begin VB.Line Line9 
         X1              =   2970
         X2              =   3060
         Y1              =   5130
         Y2              =   5130
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FF0000&
         Height          =   780
         Left            =   3150
         Top             =   135
         Width           =   6585
      End
      Begin VB.Line Line10 
         X1              =   2655
         X2              =   2970
         Y1              =   1665
         Y2              =   1665
      End
      Begin VB.Line Line11 
         X1              =   2880
         X2              =   2880
         Y1              =   315
         Y2              =   1575
      End
      Begin VB.Line Line12 
         X1              =   2880
         X2              =   3060
         Y1              =   315
         Y2              =   315
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "窗体控制"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3285
         MouseIcon       =   "Help2.frx":3930
         TabIndex        =   13
         Top             =   225
         Width           =   720
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FF0000&
         Height          =   375
         Left            =   3375
         Top             =   1035
         Width           =   6360
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "窗体控制"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3510
         MouseIcon       =   "Help2.frx":3A82
         TabIndex        =   12
         Top             =   1125
         Width           =   720
      End
      Begin VB.Line Line13 
         X1              =   3150
         X2              =   3150
         Y1              =   1170
         Y2              =   1755
      End
      Begin VB.Line Line14 
         X1              =   3150
         X2              =   3285
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Line Line15 
         X1              =   5625
         X2              =   5625
         Y1              =   1485
         Y2              =   1755
      End
      Begin VB.Line Line16 
         X1              =   5625
         X2              =   6750
         Y1              =   1485
         Y2              =   1485
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FF0000&
         Height          =   960
         Left            =   6930
         Top             =   1530
         Width           =   2805
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "媒体按钮"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   7065
         MouseIcon       =   "Help2.frx":3BD4
         TabIndex        =   11
         Top             =   1620
         Width           =   720
      End
      Begin VB.Line Line17 
         X1              =   5985
         X2              =   5985
         Y1              =   1530
         Y2              =   1755
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "播放顺序"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   6975
         MouseIcon       =   "Help2.frx":3D26
         TabIndex        =   10
         Top             =   2700
         Width           =   720
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "用于控制播放机按原序或是按随机顺序播放媒体列表中的曲目。"
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   6975
         TabIndex        =   9
         Top             =   2925
         Width           =   2655
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FF0000&
         Height          =   780
         Left            =   6840
         Top             =   2610
         Width           =   2895
      End
      Begin VB.Line Line18 
         X1              =   6840
         X2              =   6750
         Y1              =   1665
         Y2              =   1665
      End
      Begin VB.Line Line19 
         X1              =   6570
         X2              =   6570
         Y1              =   1800
         Y2              =   3645
      End
      Begin VB.Line Line20 
         X1              =   6750
         X2              =   6750
         Y1              =   1485
         Y2              =   1665
      End
      Begin VB.Line Line21 
         X1              =   5985
         X2              =   6660
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Line Line22 
         X1              =   6660
         X2              =   6660
         Y1              =   1530
         Y2              =   2745
      End
      Begin VB.Line Line23 
         X1              =   6660
         X2              =   6750
         Y1              =   2745
         Y2              =   2745
      End
      Begin VB.Line Line24 
         X1              =   3195
         X2              =   3195
         Y1              =   3690
         Y2              =   4365
      End
      Begin VB.Line Line25 
         X1              =   6345
         X2              =   6570
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Shape Shape14 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FF0000&
         Height          =   960
         Left            =   6750
         Top             =   3510
         Width           =   2985
      End
      Begin VB.Line Line26 
         X1              =   6030
         X2              =   6120
         Y1              =   5625
         Y2              =   5625
      End
      Begin VB.Line Line27 
         X1              =   6570
         X2              =   6660
         Y1              =   3645
         Y2              =   3645
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "循环控制"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   6885
         MouseIcon       =   "Help2.frx":3E78
         TabIndex        =   8
         Top             =   3600
         Width           =   720
      End
      Begin VB.Line Line29 
         X1              =   6030
         X2              =   6030
         Y1              =   3285
         Y2              =   5625
      End
      Begin VB.Shape Shape15 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FF0000&
         Height          =   780
         Left            =   6570
         Top             =   4590
         Width           =   3165
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "功能菜单"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   6705
         MouseIcon       =   "Help2.frx":3FCA
         TabIndex        =   7
         Top             =   4680
         Width           =   720
      End
      Begin VB.Line Line31 
         X1              =   6390
         X2              =   6480
         Y1              =   4725
         Y2              =   4725
      End
      Begin VB.Line Line32 
         X1              =   6390
         X2              =   6390
         Y1              =   3735
         Y2              =   4725
      End
      Begin VB.Shape Shape16 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FF0000&
         Height          =   330
         Left            =   6210
         Top             =   5490
         Width           =   3525
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "曲目列表"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   6390
         MouseIcon       =   "Help2.frx":411C
         TabIndex        =   6
         Top             =   5580
         Width           =   720
      End
      Begin VB.Shape Shape17 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00FF0000&
         Height          =   1140
         Left            =   3375
         Top             =   4230
         Width           =   2445
      End
      Begin VB.Label Label27 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "曲目列表,支持拖放文件。"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7245
         TabIndex        =   27
         Top             =   5580
         Width           =   2265
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "用于控制播放机是否循环播放媒体列表中的曲目,处于循环状态时按钮呈高亮否则色态暗淡。"
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   6885
         TabIndex        =   25
         Top             =   3825
         Width           =   2715
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "在播放状态时单击显示媒体信息;在闲置状态时则弹出媒体播放向导。"
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   7065
         TabIndex        =   24
         Top             =   1845
         Width           =   2565
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "用于显示当前媒体的相关信息如艺术家、唱片集、流派、版权等。"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4365
         TabIndex        =   23
         Top             =   1125
         Width           =   5265
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "用于控制窗体大小,从左到右依次是最小化窗体、最大化窗体和关闭窗体;你可以在选项中设置最小化时的显示方式以及最大化是对视频显示的处理。"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3285
         TabIndex        =   22
         Top             =   450
         Width           =   6330
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "用于显示当前媒体的播放状态和时间信息。"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   270
         TabIndex        =   21
         Top             =   5310
         Width           =   2550
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "用于监视和定位当前媒体的播放进度;在未安装 CD 数字音频插件前对 CD 播放无效。"
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   270
         TabIndex        =   15
         Top             =   4050
         Width           =   2160
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "用于显示视频媒体;在安装音乐可视化插件后,该窗口将根据正在播放的音频波形显示图案。"
         ForeColor       =   &H80000008&
         Height          =   810
         Left            =   270
         TabIndex        =   17
         Top             =   2790
         Width           =   2160
      End
      Begin VB.Label Label51 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "用于管理和播放多媒体文件的应用程序环境;支持文件拖放让你能把媒体文件拖放到窗体上播放机将自动打开接收到的文件。"
         ForeColor       =   &H80000008&
         Height          =   990
         Left            =   270
         TabIndex        =   20
         Top             =   1350
         Width           =   2160
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   2745
         Picture         =   "Help2.frx":426E
         Top             =   1485
         Width           =   3750
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "用于控制媒体播放;从左到右依次为:前一曲、倒退、播放、暂停、停止、弹出、快进。下一曲。"
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   3510
         TabIndex        =   29
         Top             =   4545
         Width           =   2205
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "用于设定播放音量。"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3195
         TabIndex        =   31
         Top             =   5580
         Width           =   1830
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   6180
      Left            =   0
      TabIndex        =   48
      Top             =   540
      Width           =   10095
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "http://fmxz.51.net"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   2970
         MouseIcon       =   "Help2.frx":7590
         MousePointer    =   99  'Custom
         TabIndex        =   62
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Image Image2 
         Height          =   3450
         Left            =   5625
         Picture         =   "Help2.frx":76E2
         Top             =   2520
         Width           =   4500
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "leask@21cn.com"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   7020
         MouseIcon       =   "Help2.frx":8E9A
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label35 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   $"Help2.frx":8FEC
         ForeColor       =   &H00000000&
         Height          =   1230
         Left            =   1260
         TabIndex        =   47
         Top             =   1080
         Width           =   7890
      End
      Begin VB.Label Label34 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "如何联系作者以获得帮助？"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   180
         TabIndex        =   49
         Top             =   135
         Width           =   3435
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   6180
      Left            =   0
      TabIndex        =   44
      Top             =   540
      Width           =   10095
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   5325
         Left            =   1890
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         Text            =   "Help2.frx":9096
         Top             =   630
         Width           =   7980
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "通过本栏目了解 Snowman Media ilxz 所能播放的媒体格式。"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   180
         TabIndex        =   46
         Top             =   135
         Width           =   5550
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   6180
      Left            =   0
      TabIndex        =   35
      Top             =   540
      Width           =   10095
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   5325
         Left            =   1890
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Text            =   "Help2.frx":98CD
         Top             =   630
         Width           =   7980
      End
      Begin VB.Label Label36 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "通过本栏目了解 Snowman Media ilxz 的快捷键分配以实现快速操作。"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   180
         TabIndex        =   36
         Top             =   135
         Width           =   6405
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   5910
      Left            =   0
      TabIndex        =   53
      Top             =   540
      Width           =   9870
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "返回"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   8010
         MouseIcon       =   "Help2.frx":A33D
         MousePointer    =   99  'Custom
         TabIndex        =   61
         Top             =   5175
         Width           =   1185
      End
      Begin VB.Shape Shape19 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   330
         Left            =   8010
         Top             =   5085
         Width           =   1200
      End
      Begin VB.Label Label52 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "软件环境:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   4590
         TabIndex        =   57
         Top             =   540
         Width           =   810
      End
      Begin VB.Label Label55 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   $"Help2.frx":A48F
         ForeColor       =   &H80000008&
         Height          =   3030
         Left            =   1620
         TabIndex        =   60
         Top             =   495
         Width           =   3075
      End
      Begin VB.Label Label54 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "硬件环境:"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   630
         TabIndex        =   59
         Top             =   495
         Width           =   810
      End
      Begin VB.Label Label53 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   $"Help2.frx":A539
         ForeColor       =   &H80000008&
         Height          =   1860
         Left            =   5625
         TabIndex        =   58
         Top             =   540
         Width           =   4110
      End
      Begin VB.Label Label49 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "有关说明:"
         ForeColor       =   &H00C000C0&
         Height          =   180
         Left            =   630
         TabIndex        =   56
         Top             =   3690
         Width           =   810
      End
      Begin VB.Label Label48 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "通过本栏目了解运行 Snowman Media ilxz 的系统要求。"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   180
         TabIndex        =   54
         Top             =   135
         Width           =   5010
      End
      Begin VB.Label Label47 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   $"Help2.frx":A5E3
         ForeColor       =   &H80000008&
         Height          =   1590
         Left            =   1620
         TabIndex        =   55
         Top             =   3690
         Width           =   5955
      End
   End
   Begin VB.Label Label32 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "自述文件"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   7965
      MouseIcon       =   "Help2.frx":A6BB
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   180
      Width           =   720
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "许可协议"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   8910
      MouseIcon       =   "Help2.frx":A80D
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   180
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   6435
      X2              =   9765
      Y1              =   405
      Y2              =   405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "支持格式"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3870
      MouseIcon       =   "Help2.frx":A95F
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   180
      Width           =   1185
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "初步了解"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   90
      MouseIcon       =   "Help2.frx":AAB1
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   180
      Width           =   1185
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "便捷操作"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2610
      MouseIcon       =   "Help2.frx":AC03
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   180
      Width           =   1185
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "联系我们"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5130
      MouseIcon       =   "Help2.frx":AD55
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   180
      Width           =   1185
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "快速入门"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1350
      MouseIcon       =   "Help2.frx":AEA7
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   180
      Width           =   1185
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Left            =   2610
      Top             =   90
      Width           =   1200
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Left            =   90
      Top             =   90
      Width           =   1200
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Left            =   3870
      Top             =   90
      Width           =   1200
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Left            =   1350
      Top             =   90
      Width           =   1200
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Left            =   5130
      Top             =   90
      Width           =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = True
Frame5.Visible = False
Frame6.Visible = False
Shape8.BackColor = &HFFFFFF
Shape5.BackColor = &HFFFFFF
Shape3.BackColor = &HFFC0C0
Shape4.BackColor = &HFFFFFF
Shape2.BackColor = &HFFFFFF

End Sub

Private Sub Label11_Click()
Ly.OpenFile App.Path + "\许可协议.txt"

End Sub


Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Shape8.BackColor = &HFFFFFF
Shape5.BackColor = &HFFFFFF
Shape3.BackColor = &HFFFFFF
Shape4.BackColor = &HFFC0C0
Shape2.BackColor = &HFFFFFF

End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = True
Frame6.Visible = False
Shape8.BackColor = &HFFFFFF
Shape5.BackColor = &HFFC0C0
Shape3.BackColor = &HFFFFFF
Shape4.BackColor = &HFFFFFF
Shape2.BackColor = &HFFFFFF

End Sub

Private Sub Label32_Click()
Ly.OpenFile App.Path + "\自述文件.txt"
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Shape8.BackColor = &HFFC0C0
Shape5.BackColor = &HFFFFFF
Shape3.BackColor = &HFFFFFF
Shape4.BackColor = &HFFFFFF
Shape2.BackColor = &HFFFFFF
End Sub

Private Sub Label43_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = True

End Sub

Private Sub Label44_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False

End Sub

Private Sub Label45_Click()
Ly.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "NetShow", True
Ly.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "NetFile", "http://www.51.net"
Shell App.Path + "\SmM_IntBrowser.exe", vbMinimizedFocus

End Sub

Private Sub Label57_Click()
Ly.SendMail ("leask@21cn.com")
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
Shape8.BackColor = &HFFFFFF
Shape5.BackColor = &HFFFFFF
Shape3.BackColor = &HFFFFFF
Shape4.BackColor = &HFFFFFF
Shape2.BackColor = &HFFC0C0

End Sub
