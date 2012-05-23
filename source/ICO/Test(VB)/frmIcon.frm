VERSION 5.00
Object = "{CFB094A6-8FF0-4EF7-A644-ED122CC38E57}#1.0#0"; "TASKICON.OCX"
Begin VB.Form frmIcon 
   Caption         =   "TaskIcon ActiveX Demo"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Animation"
      Height          =   495
      Left            =   3120
      TabIndex        =   17
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hidden Demo"
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      Top             =   960
      Width           =   1335
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "Hidden TaskIcon"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Text            =   "TaskIcon ActiveX"
      Top             =   600
      Width           =   2415
   End
   Begin VB.ListBox lstEvent 
      Height          =   1035
      Index           =   1
      Left            =   2400
      TabIndex        =   10
      Top             =   3480
      Width           =   2055
   End
   Begin VB.ListBox lstEvent 
      Height          =   1035
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Show Infomation Message for Win2000"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4335
      Begin VB.CommandButton Command1 
         Caption         =   "Show"
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton optIcon 
         Caption         =   "Error"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optIcon 
         Caption         =   "Warning"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton optIcon 
         Caption         =   "Info"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optIcon 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   1215
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmIcon.frx":0000
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "TaskIcon ..."
      Top             =   120
      Width           =   2415
   End
   Begin TASKICONLib.TaskIcon TaskIcon 
      Left            =   3960
      Top             =   120
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ToolTipText     =   "TaskIcon ..."
      TitleText       =   "TaskIcon ActiveX"
      TitleBackColor  =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   7
      Left            =   720
      Picture         =   "frmIcon.frx":0037
      Top             =   5280
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   6
      Left            =   120
      Picture         =   "frmIcon.frx":0479
      Top             =   5280
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   5
      Left            =   3120
      Picture         =   "frmIcon.frx":08BB
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   4
      Left            =   2520
      Picture         =   "frmIcon.frx":0CFD
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   3
      Left            =   1920
      Picture         =   "frmIcon.frx":113F
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   2
      Left            =   1320
      Picture         =   "frmIcon.frx":1581
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   720
      Picture         =   "frmIcon.frx":19C3
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmIcon.frx":1E05
      Top             =   4680
      Width           =   480
   End
   Begin VB.Label Label4 
      Caption         =   "Menu Title"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Menu Event"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Icon Event"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ToolTipText"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nIcon As Integer



Private Sub chkHidden_Click()
    If chkHidden.Value = 1 Then
        TaskIcon.Hidden = True
    Else
        TaskIcon.Hidden = False
    End If
End Sub


Private Sub Command1_Click()
    TaskIcon.ShowMsg Text2.Text, nIcon, "TaskIcon", 10
End Sub

Private Sub Command2_Click()
    Me.Visible = False
End Sub

Private Sub Command3_Click()
    Static entry As Boolean
    Dim t As Double, i As Integer, c As Integer
    If entry Then Exit Sub
    
    entry = True
    For c = 0 To 5
        For i = 0 To imgIcon.Count - 1
            TaskIcon.Icon = imgIcon(i).Picture
            t = Timer
            Do While (Timer - t) < 0.1
                DoEvents
            Loop
        Next i
    Next c
    entry = False
End Sub

Private Sub Form_Load()
    Dim id As Integer
    
    For id = 1 To 10
        Select Case id
            Case 5
                TaskIcon.AddMenuItem MF_SEPARATOR, id, ""
            Case 6
                TaskIcon.AddMenuItem MF_STRING, id, "Check Item" & id
            Case Else
            TaskIcon.AddMenuItem MF_STRING, id, "TaskIcon Menu Test" & id
        End Select
    Next id
    TaskIcon.AddMenuItem MF_SEPARATOR, 0, ""
    TaskIcon.AddMenuItem MF_STRING, id, "Show Demo Software"
    TaskIcon.EnableMenuItem(3) = False
    TaskIcon.CheckMenuItem(6) = True
    
    TaskIcon.ShowIcon
End Sub


Private Sub optIcon_Click(Index As Integer)
    nIcon = Index
End Sub

Private Sub TaskIcon_OnIconEvent(ByVal EventMsg As Long)
    lstEvent(0).AddItem EventMsg
End Sub

Private Sub TaskIcon_OnMenuEvent(ByVal id As Integer)
    lstEvent(1).AddItem id
    
    If id = 6 Then
        TaskIcon.CheckMenuItem(6) = Not TaskIcon.CheckMenuItem(6)
    End If
    
    If id = 11 Then
        Me.Visible = True
    End If
End Sub


Private Sub Text1_Change()
    TaskIcon.ToolTipText = Text1.Text
End Sub


Private Sub Text3_Change()
    TaskIcon.TitleText = Text3.Text
End Sub


