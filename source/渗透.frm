VERSION 5.00
Begin VB.Form frmColourise 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "渗透颜色  － Snowman Media P.B  1.0"
   ClientHeight    =   1200
   ClientLeft      =   5220
   ClientTop       =   4275
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "渗透.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraSep 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      TabIndex        =   10
      Top             =   765
      Width           =   8475
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3915
      ScaleHeight     =   255
      ScaleWidth      =   885
      TabIndex        =   7
      Top             =   45
      Width           =   915
      Begin VB.Label lblSelected 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   "选取:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   45
         TabIndex        =   9
         Top             =   45
         Width           =   450
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   135
      ScaleHeight     =   255
      ScaleWidth      =   3630
      TabIndex        =   6
      Top             =   45
      Width           =   3660
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   "单击你想要在图片上渗透的颜色:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   45
         TabIndex        =   8
         Top             =   45
         Width           =   2610
      End
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3735
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   4
      Top             =   855
      Width           =   1095
      Begin VB.Label Label14 
         BackColor       =   &H0000FFFF&
         Caption         =   "取消(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   45
         TabIndex        =   5
         Top             =   45
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2565
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   2
      Top             =   855
      Width           =   1095
      Begin VB.Label Label13 
         BackColor       =   &H0000FFFF&
         Caption         =   "确定(&Y)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   45
         TabIndex        =   3
         Top             =   45
         Width           =   1680
      End
   End
   Begin VB.PictureBox picSelected 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3915
      ScaleHeight     =   255
      ScaleWidth      =   885
      TabIndex        =   1
      Top             =   360
      Width           =   915
   End
   Begin VB.PictureBox picHue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   135
      ScaleHeight     =   255
      ScaleWidth      =   3630
      TabIndex        =   0
      Top             =   360
      Width           =   3660
   End
End
Attribute VB_Name = "frmColourise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bCancel As Boolean
Private m_fHue As Single

Public Property Get Cancelled() As Boolean
   Cancelled = m_bCancel
End Property

Public Property Get Hue() As Single
   Hue = m_fHue
End Property

Private Sub cmdCancel_Click()
 
End Sub

Private Sub cmdOK_Click()
 
End Sub

Private Sub Form_Load()
Dim h As Single
Dim r As Long, g As Long, b As Long
Dim lH As Long
Dim x As Long
   m_bCancel = True
   lH = picHue.ScaleHeight
   For h = -40 To 200
      HLSToRGB h / 40, 1, 0.5, r, g, b
      picHue.Line (x, 0)-(x + Screen.TwipsPerPixelX, lH), RGB(r, g, b), BF
      x = x + Screen.TwipsPerPixelX
   Next h
   picHue.Refresh
   picHue_MouseDown 1, 0, 0, 0
End Sub

Private Sub Label13_Click()
  m_bCancel = False
   Unload Me
End Sub

Private Sub Label14_Click()
  Unload Me
End Sub

Private Sub picHue_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim r As Long, g As Long, b As Long
   If (x > 0) And (y > 0) And (x < picHue.ScaleWidth) And (y < picHue.ScaleHeight) Then
      m_fHue = ((x \ Screen.TwipsPerPixelX) - 40) / 40
      HLSToRGB m_fHue, 1, 0.5, r, g, b
      picSelected.BackColor = RGB(r, g, b)
   End If

End Sub

Private Sub picHue_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If (Button And vbLeftButton) = vbLeftButton Then
      picHue_MouseDown Button, Shift, x, y
   End If

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



