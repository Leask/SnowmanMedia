VERSION 5.00
Begin VB.Form frmAddNoise 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "添加杂点  － Snowman Media P.B.  1.0"
   ClientHeight    =   750
   ClientLeft      =   4770
   ClientTop       =   2415
   ClientWidth     =   4785
   Icon            =   "杂点.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.OptionButton optType 
      Caption         =   "随机添加(&N)"
      Height          =   240
      Index           =   1
      Left            =   225
      TabIndex        =   7
      Top             =   405
      Width           =   1680
   End
   Begin VB.OptionButton optType 
      Caption         =   "一至添加(&S)"
      Height          =   240
      Index           =   0
      Left            =   225
      TabIndex        =   8
      Top             =   90
      Value           =   -1  'True
      Width           =   1680
   End
   Begin VB.Frame fraSep 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   2295
      TabIndex        =   6
      Top             =   360
      Width           =   8565
   End
   Begin VB.TextBox txtAmount 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3375
      TabIndex        =   4
      Text            =   "20"
      Top             =   45
      Width           =   645
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2295
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   2
      Top             =   405
      Width           =   1095
      Begin VB.Label Label13 
         BackColor       =   &H0000FFFF&
         Caption         =   "确定(&Y)"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   45
         TabIndex        =   3
         Top             =   45
         Width           =   1680
      End
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3555
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   0
      Top             =   405
      Width           =   1095
      Begin VB.Label Label14 
         BackColor       =   &H0000FFFF&
         Caption         =   "取消(&C)"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   45
         TabIndex        =   1
         Top             =   45
         Width           =   1635
      End
   End
   Begin VB.Label Label1 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4140
      TabIndex        =   5
      Top             =   90
      Width           =   375
   End
   Begin VB.Label lblAmount 
      AutoSize        =   -1  'True
      Caption         =   "百分比(&P):"
      Height          =   240
      Left            =   2295
      TabIndex        =   9
      Top             =   90
      Width           =   2040
   End
End
Attribute VB_Name = "frmAddNoise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bCancelled As Boolean
Private m_bRandom As Boolean
Private m_lPercent As Long

Public Property Get Cancelled() As Boolean
    Cancelled = m_bCancelled
End Property

Public Property Get Random() As Boolean
    Random = m_bRandom
End Property
Public Property Get Percentage() As Long
    Percentage = m_lPercent
End Property




Private Sub cmdCancel_Click()

End Sub

Private Sub Form_Load()
    m_bCancelled = True
End Sub

Private Sub Label13_Click()
    m_lPercent = CLng(txtAmount.Text)
    m_bCancelled = False
    Unload Me
End Sub

Private Sub Label14_Click()
Unload Me
End Sub

Private Sub optType_Click(Index As Integer)
    m_bRandom = optType(1).Value
End Sub
Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label13.BackColor = &HFF0000
Picture9.BackColor = &HFF0000
Label13.ForeColor = &HFFFF&
End Sub
Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label14.BackColor = &HFF0000
Picture10.BackColor = &HFF0000
Label14.ForeColor = &HFFFF&
End Sub
Private Sub form_mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture9.BackColor = &HFFFF&

Picture10.BackColor = &HFFFF&
Label13.ForeColor = &HFF0000
Label13.BackColor = &HFFFF&
Label14.ForeColor = &HFF0000
Label14.BackColor = &HFFFF&
End Sub


