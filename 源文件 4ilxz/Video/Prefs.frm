VERSION 5.00
Begin VB.Form frmPrefs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ѡ��"
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
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&Y)"
      Height          =   330
      Left            =   2205
      TabIndex        =   12
      Top             =   3735
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(C)"
      Height          =   330
      Left            =   3645
      TabIndex        =   11
      Top             =   3735
      Width           =   1185
   End
   Begin VB.Frame staticframe 
      Caption         =   "��Ƶ����Ƶͬ��:"
      Height          =   1095
      Index           =   2
      Left            =   270
      TabIndex        =   2
      Top             =   2385
      Width           =   4695
      Begin VB.OptionButton optSynch 
         Caption         =   "����(&N)               (�����ȿ��ܲ�һ��)"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   10
         Top             =   585
         Width           =   4290
      End
      Begin VB.OptionButton optSynch 
         Caption         =   "ͬ����Ƶ����Ƶ(&V)     (��Ƶ֡�ʿ��ܻ�ı�)"
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
      Caption         =   "���֡��:"
      Height          =   1050
      Index           =   1
      Left            =   270
      TabIndex        =   1
      Top             =   1170
      Width           =   4695
      Begin VB.OptionButton optMaxFrames 
         Caption         =   "324,000    (3 Сʱ;ÿ�� 30 ֡)"
         Height          =   240
         Index           =   1
         Left            =   270
         TabIndex        =   8
         Top             =   585
         Width           =   3615
      End
      Begin VB.OptionButton optMaxFrames 
         Caption         =   "27,000    (15 ����;ÿ�� 30 ֡)"
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
      Caption         =   "����ɫ:"
      Height          =   780
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   225
      Width           =   4695
      Begin VB.OptionButton optColor 
         Caption         =   "��ɫ"
         Height          =   285
         Index           =   3
         Left            =   3240
         TabIndex        =   6
         Top             =   270
         Width           =   960
      End
      Begin VB.OptionButton optColor 
         Caption         =   "���ɫ"
         Height          =   285
         Index           =   2
         Left            =   2115
         TabIndex        =   5
         Top             =   270
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.OptionButton optColor 
         Caption         =   "ǳ��ɫ"
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   4
         Top             =   270
         Width           =   960
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Ĭ��"
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
    Call SaveSetting(App.Title, "ѡ��", "����ɫ", color)
    'commit maxframes
    Call SaveSetting(App.Title, "ѡ��", "���֡��", maxframes)
    'commit streammaster
    Call SaveSetting(App.Title, "ѡ��", "������", streammaster)
    
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Index As Long
    
    'load current settings
    color = Val(GetSetting(App.Title, "ѡ��", "����ɫ", "&H404040"))
    maxframes = Val(GetSetting(App.Title, "ѡ��", "���֡��", INDEX_15_MINUTES))
    streammaster = GetSetting(App.Title, "ѡ��", "������", "��Ƶ")
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
    If Trim$(UCase(streammaster)) = "��Ƶ" Then
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
            streammaster = "��Ƶ"
        Case 1 ' None
            streammaster = "��"
    End Select
End Sub
