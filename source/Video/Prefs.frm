VERSION 5.00
Begin VB.Form frmPrefs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选项"
   ClientHeight    =   4260
   ClientLeft      =   1770
   ClientTop       =   3075
   ClientWidth     =   5265
   Icon            =   "Prefs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&Y)"
      Height          =   330
      Left            =   2205
      TabIndex        =   12
      Top             =   3735
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(C)"
      Height          =   330
      Left            =   3645
      TabIndex        =   11
      Top             =   3735
      Width           =   1185
   End
   Begin VB.Frame staticframe 
      Caption         =   "视频和音频同步:"
      Height          =   1095
      Index           =   2
      Left            =   270
      TabIndex        =   2
      Top             =   2385
      Width           =   4695
      Begin VB.OptionButton optSynch 
         Caption         =   "忽略(&N)               (流长度可能不一致)"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   10
         Top             =   585
         Width           =   4290
      End
      Begin VB.OptionButton optSynch 
         Caption         =   "同步视频到音频(&V)     (视频帧率可能会改变)"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   9
         Top             =   270
         Value           =   -1  'True
         Width           =   4290
      End
   End
   Begin VB.Frame staticframe 
      Caption         =   "最大帧数:"
      Height          =   1050
      Index           =   1
      Left            =   270
      TabIndex        =   1
      Top             =   1170
      Width           =   4695
      Begin VB.OptionButton optMaxFrames 
         Caption         =   "324,000    (3 小时;每秒 30 帧)"
         Height          =   240
         Index           =   1
         Left            =   270
         TabIndex        =   8
         Top             =   585
         Width           =   3615
      End
      Begin VB.OptionButton optMaxFrames 
         Caption         =   "27,000    (15 分钟;每秒 30 帧)"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   7
         Top             =   270
         Value           =   -1  'True
         Width           =   3435
      End
   End
   Begin VB.Frame staticframe 
      Caption         =   "背景色:"
      Height          =   780
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   225
      Width           =   4695
      Begin VB.OptionButton optColor 
         Caption         =   "黑色"
         Height          =   285
         Index           =   3
         Left            =   3240
         TabIndex        =   6
         Top             =   270
         Width           =   960
      End
      Begin VB.OptionButton optColor 
         Caption         =   "深灰色"
         Height          =   285
         Index           =   2
         Left            =   2115
         TabIndex        =   5
         Top             =   270
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton optColor 
         Caption         =   "浅灰色"
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   4
         Top             =   270
         Width           =   960
      End
      Begin VB.OptionButton optColor 
         Caption         =   "默认"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   3
         Top             =   270
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private color As Long
Private maxframes As Long
Private streammaster As String

Private Sub cmdCancel_Click()
    'lose all changes
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'commit back color changes
    frmMain.BackColor = color
    Call SaveSetting(App.Title, "选项", "背景色", color)
    'commit maxframes
    Call SaveSetting(App.Title, "选项", "最大帧数", maxframes)
    'commit streammaster
    Call SaveSetting(App.Title, "选项", "流主管", streammaster)
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Index As Long
    
    'load current settings
    color = Val(GetSetting(App.Title, "选项", "背景色", "&H404040"))
    maxframes = Val(GetSetting(App.Title, "选项", "最大帧数", INDEX_15_MINUTES))
    streammaster = GetSetting(App.Title, "选项", "流主管", "音频")
    'and set up form
    Select Case color
        Case &H0& 'black
            Index = 3
        Case &HC0C0C0 'lt gray
            Index = 1
        Case &H404040 'dk gray
            Index = 2
        Case &H80000005 'default window
            Index = 0
        Case Else
            Index = 0
    End Select
    optColor(Index).Value = True 'set correct color option
    If maxframes <> INDEX_15_MINUTES Then
        Index = 1
    Else
        Index = 0
    End If
    optMaxFrames(Index).Value = True 'set correct frames option
    If Trim$(UCase(streammaster)) = "音频" Then
        Index = 0
    Else
        Index = 1
    End If
    optSynch(Index).Value = True
    
End Sub

Private Sub optColor_Click(Index As Integer)
    Select Case Index
        Case 0 'default
            color = &H80000005  'use windows' system
        Case 1 'lt gray
            color = &HC0C0C0
        Case 2 'dk gray
            color = &H404040
        Case 3 'black
            color = &H0&
    End Select
End Sub

Private Sub optMaxFrames_Click(Index As Integer)
    Select Case Index
        Case 0 '27000 (15 minutes)
            maxframes = INDEX_15_MINUTES
        Case 1 '324000 (3 hours)
            maxframes = INDEX_3_HOURS
    End Select
End Sub

Private Sub optSynch_Click(Index As Integer)
    Select Case Index
        Case 0 'Audio
            streammaster = "音频"
        Case 1 ' None
            streammaster = "无"
    End Select
End Sub
