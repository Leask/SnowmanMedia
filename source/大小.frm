VERSION 5.00
Begin VB.Form frmNewSize 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "改变大小  - Snowman Media P.B.  1.0"
   ClientHeight    =   1275
   ClientLeft      =   5340
   ClientTop       =   2520
   ClientWidth     =   5280
   Icon            =   "大小.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4050
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   15
      Top             =   900
      Width           =   1095
      Begin VB.Label Label14 
         BackColor       =   &H0000FFFF&
         Caption         =   "取消(&C)"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   45
         TabIndex        =   16
         Top             =   45
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2880
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   13
      Top             =   900
      Width           =   1095
      Begin VB.Label Label13 
         BackColor       =   &H0000FFFF&
         Caption         =   "确定(&Y)"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   45
         TabIndex        =   14
         Top             =   45
         Width           =   1680
      End
   End
   Begin VB.Frame fraSep 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   15
      Left            =   2610
      TabIndex        =   10
      Top             =   810
      Width           =   2535
   End
   Begin VB.TextBox txtWidth 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1035
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   360
      Width           =   600
   End
   Begin VB.TextBox txtHeight 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1035
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   675
      Width           =   600
   End
   Begin VB.TextBox txtPercent 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4095
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "100"
      Top             =   405
      Width           =   510
   End
   Begin VB.CheckBox chkKeepProportion 
      Caption         =   "保持宽高比(&K):"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   405
      TabIndex        =   3
      Top             =   990
      Width           =   1905
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "自定义:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   2
      Top             =   45
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.OptionButton optStyle 
      Caption         =   "百分比(&P):"
      Height          =   195
      Index           =   1
      Left            =   2610
      TabIndex        =   1
      Top             =   45
      Width           =   2535
   End
   Begin VB.Label lblOrigWidth 
      Height          =   240
      Left            =   1755
      TabIndex        =   12
      Top             =   360
      Width           =   555
   End
   Begin VB.Label lblOrigHeight 
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1755
      TabIndex        =   11
      Top             =   675
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "%"
      Height          =   195
      Left            =   4770
      TabIndex        =   9
      Top             =   405
      Width           =   195
   End
   Begin VB.Label lbl 
      Caption         =   "宽(&W):"
      Height          =   240
      Left            =   405
      TabIndex        =   8
      Top             =   360
      Width           =   600
   End
   Begin VB.Label lblPercentage 
      Caption         =   "百分比(&E):"
      Height          =   195
      Left            =   2880
      TabIndex        =   7
      Top             =   405
      Width           =   2265
   End
   Begin VB.Label lblHeight 
      Caption         =   "高(&H):"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   405
      TabIndex        =   0
      Top             =   675
      Width           =   600
   End
End
Attribute VB_Name = "frmNewSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bCancelled As Boolean
Private m_lWIdth As Long
Private m_lHeight As Long
Private m_lOrigWidth As Long
Private m_lOrigHeight As Long

Public Sub SetSize(ByVal lW As Long, ByVal lH As Long)
    m_lOrigWidth = lW
    m_lOrigHeight = lH
End Sub

Public Property Get ImageWidth() As Long
    ImageWidth = m_lWIdth
End Property
Public Property Get ImageHeight() As Long
    ImageHeight = m_lHeight
End Property

Public Property Get Cancelled() As Boolean
    Cancelled = m_bCancelled
End Property

Private Sub EnableControls()
Dim bS As Boolean
Dim lC1 As Long, lC2 As Long
    lC2 = vbButtonFace: lC1 = vbWindowBackground
    bS = (optStyle(0).Value)
    txtWidth.Locked = Not (bS)
    txtWidth.BackColor = lC2
    txtHeight.Locked = Not (bS)
    txtHeight.BackColor = lC2
    chkKeepProportion.Enabled = bS
    txtPercent.Locked = bS
    txtPercent.BackColor = lC1
End Sub

Private Sub cmdCancel_Click()
   
End Sub

Private Sub cmdOK_Click()
 
End Sub

Private Sub Form_Load()
    m_bCancelled = True
    lblOrigWidth.Caption = m_lOrigWidth
    lblOrigHeight.Caption = m_lOrigHeight
    txtWidth.Text = m_lOrigWidth
    txtHeight.Text = m_lOrigHeight
End Sub

Private Sub Label13_Click()
   m_lWIdth = CLng(txtWidth.Text)
    m_lHeight = CLng(txtHeight.Text)
    m_bCancelled = False
    Unload Me
End Sub

Private Sub Label14_Click()
 Unload Me
End Sub

Private Sub optStyle_Click(Index As Integer)
    EnableControls
End Sub

Private Sub txtHeight_Change()
Dim l As Long
    If (chkKeepProportion.Value = Checked) Then
        If (txtHeight.Tag = "") Then
            If IsNumeric(txtHeight) Then
                l = CLng(txtHeight)
                txtWidth.Tag = "txtHeight"
                txtWidth.Text = l * CLng(lblOrigWidth.Caption) \ CLng(lblOrigHeight.Caption)
                txtWidth.Tag = ""
            End If
        End If
    End If

End Sub

Private Sub txtPercent_Change()
Dim l As Double
    If (optStyle(1).Value) Then
        If IsNumeric(txtPercent) Then
            l = CLng(txtPercent) / 100
            txtWidth.Text = CLng(CLng(lblOrigWidth.Caption) * l)
            txtHeight.Text = CLng(CLng(lblOrigHeight.Caption) * l)
        End If
    End If
End Sub

Private Sub txtWidth_Change()
Dim l As Long
    If (chkKeepProportion.Value = Checked) Then
        If txtWidth.Tag = "" Then
            If IsNumeric(txtWidth) Then
                l = CLng(txtWidth)
                txtHeight.Tag = "txtWidth"
                txtHeight.Text = l * CLng(lblOrigHeight.Caption) \ CLng(lblOrigWidth.Caption)
                txtHeight.Tag = ""
            End If
        End If
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
