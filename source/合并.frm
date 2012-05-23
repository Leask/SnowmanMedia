VERSION 5.00
Begin VB.Form frmCombination 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "合并图片  - Snowman Media Pictures Browser  1.0"
   ClientHeight    =   2520
   ClientLeft      =   6300
   ClientTop       =   4215
   ClientWidth     =   6450
   Icon            =   "合并.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   1575
      ScaleHeight     =   570
      ScaleWidth      =   4665
      TabIndex        =   17
      Top             =   1260
      Width           =   4695
      Begin VB.OptionButton optCombine 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "明亮(&L)"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   2
         Left            =   45
         MaskColor       =   &H00FF0000&
         TabIndex        =   23
         Top             =   315
         Width           =   4695
      End
      Begin VB.OptionButton optCombine 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "暗淡(&D)"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   45
         MaskColor       =   &H00FF0000&
         TabIndex        =   22
         Top             =   45
         Width           =   4695
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   180
      ScaleHeight     =   255
      ScaleWidth      =   1155
      TabIndex        =   16
      Top             =   1260
      Width           =   1185
      Begin VB.Label Label1 
         BackColor       =   &H0080FFFF&
         Caption         =   "明暗(&L):"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   45
         TabIndex        =   21
         Top             =   45
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   180
      ScaleHeight     =   255
      ScaleWidth      =   1155
      TabIndex        =   15
      Top             =   855
      Width           =   1185
      Begin VB.Label lblMultiplier 
         BackColor       =   &H0080FFFF&
         Caption         =   "增强(&M):"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   20
         Tag             =   "Add"
         Top             =   45
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   180
      ScaleHeight     =   255
      ScaleWidth      =   1155
      TabIndex        =   14
      Top             =   450
      Width           =   1185
      Begin VB.Label lblOffset 
         BackColor       =   &H0080FFFF&
         Caption         =   "抵消(&O):"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   19
         Tag             =   "Add"
         Top             =   45
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   180
      ScaleHeight     =   255
      ScaleWidth      =   1155
      TabIndex        =   13
      Top             =   45
      Width           =   1185
      Begin VB.Label lblSource 
         BackColor       =   &H0080FFFF&
         Caption         =   "图象来源(&S):"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   18
         Top             =   45
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5175
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   11
      Top             =   2160
      Width           =   1095
      Begin VB.Label Label14 
         BackColor       =   &H0000FFFF&
         Caption         =   "取消(&C)"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   45
         TabIndex        =   12
         Top             =   45
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4005
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
      Begin VB.Label Label13 
         BackColor       =   &H0000FFFF&
         Caption         =   "确定(&Y)"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   45
         TabIndex        =   10
         Top             =   45
         Width           =   1680
      End
   End
   Begin VB.OptionButton optCombine 
      Caption         =   "&Add Images"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   4860
      Value           =   -1  'True
      Width           =   4830
   End
   Begin VB.Frame fraSep 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   -1170
      TabIndex        =   7
      Top             =   2025
      Width           =   8475
   End
   Begin VB.TextBox txtMultiplier 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   3960
      TabIndex        =   5
      Tag             =   "Add"
      Text            =   "1"
      Top             =   855
      Width           =   2310
   End
   Begin VB.TextBox txtOffset 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   3960
      TabIndex        =   4
      Tag             =   "Add"
      Text            =   "0"
      Top             =   450
      Width           =   2310
   End
   Begin VB.TextBox txtMultiplier 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   1575
      TabIndex        =   3
      Tag             =   "Add"
      Text            =   "1"
      Top             =   855
      Width           =   2310
   End
   Begin VB.TextBox txtOffset 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   1575
      TabIndex        =   2
      Tag             =   "Add"
      Text            =   "0"
      Top             =   450
      Width           =   2310
   End
   Begin VB.ComboBox cboSource 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   1
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   45
      Width           =   2310
   End
   Begin VB.ComboBox cboSource 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   0
      Left            =   1575
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   45
      Width           =   2310
   End
   Begin VB.Label lblInfo 
      Height          =   285
      Left            =   540
      TabIndex        =   6
      Top             =   5670
      Width           =   4830
   End
End
Attribute VB_Name = "frmCombination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bCancel As Boolean
Private m_iOpt As Long
Private m_iOffset(1 To 2) As Long
Private m_iMultiplier(1 To 2) As Long
Private m_lNewImageWidth As Long
Private m_lNewImageHeight As Long
Private m_lImageSource(1 To 2) As Long

Public Enum ECombinationTypeConstants
   eAdd = 1
   eLightest = 2
   eDarkest = 3
End Enum

Public Property Get CombinationType() As ECombinationTypeConstants
   CombinationType = m_iOpt
End Property
Public Property Get Offset(ByVal lIndex As Long)
   Offset = m_iOffset(lIndex)
End Property
Public Property Get Multiplier(ByVal lIndex As Long)
   Multiplier = m_iMultiplier(lIndex)
End Property

Public Property Get Cancelled() As Boolean
   Cancelled = m_bCancel
End Property

Public Property Get ImageSource(ByVal lIndex As Long) As Long
   ImageSource = m_lImageSource(lIndex)
End Property

Public Property Get NewImageWidth() As Long
   NewImageWidth = m_lNewImageWidth
End Property
Public Property Get NewImageHeight() As Long
   NewImageHeight = m_lNewImageHeight
End Property


Private Sub cboSource_Click(Index As Integer)
Dim lW1 As Long, lH1 As Long
Dim lW2 As Long, lH2 As Long
   If (cboSource(Index).Tag = "") Then
      lW1 = Forms(cboSource(0).ItemData(cboSource(0).ListIndex)).ImageWidth
      lH1 = Forms(cboSource(0).ItemData(cboSource(0).ListIndex)).ImageHeight
      lW2 = Forms(cboSource(1).ItemData(cboSource(1).ListIndex)).ImageWidth
      lH2 = Forms(cboSource(1).ItemData(cboSource(1).ListIndex)).ImageHeight
      If (lW1 > lW2) Then
         m_lNewImageWidth = lW2
      Else
         m_lNewImageWidth = lW1
      End If
      If (lH1 > lH2) Then
         m_lNewImageHeight = lH2
      Else
         m_lNewImageHeight = lH1
      End If
      lblInfo.Caption = "输出图片大小: " & m_lNewImageWidth & " x " & m_lNewImageHeight
   End If
   m_lImageSource(Index + 1) = cboSource(Index).ItemData(cboSource(Index).ListIndex)
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()

End Sub

Private Sub Form_Load()
Dim iFrm As Long
Dim sItem As String
Dim i As Long

   m_bCancel = True
   For iFrm = 0 To Forms.Count - 1
      If Forms(iFrm).Name = "frmImage" Then
         sItem = Forms(iFrm).FileTitle
         With cboSource(0)
            .AddItem sItem
            .ItemData(.NewIndex) = iFrm
         End With
         With cboSource(1)
            .AddItem sItem
            .ItemData(.NewIndex) = iFrm
         End With
      End If
   Next iFrm
   cboSource(0).Tag = "DONT"
   cboSource(0).ListIndex = 0
   cboSource(0).Tag = ""
   cboSource(1).ListIndex = Abs(cboSource(1).ListCount > 1)
End Sub

Private Sub Label13_Click()
On Error GoTo ValFail
   m_iOpt = -1 * (optCombine(0).Value + 2 * optCombine(1).Value + 3 * optCombine(2).Value)
   If (m_iOpt > 0) Then
      m_iOffset(1) = CLng(txtOffset(0))
      m_iOffset(2) = CLng(txtOffset(1))
      m_iMultiplier(1) = CLng(txtMultiplier(0))
      m_iMultiplier(2) = CLng(txtMultiplier(1))
   End If
   Unload Me
   m_bCancel = False
   Exit Sub

ValFail:
   MsgBox "发生错误.", vbInformation
   Exit Sub
End Sub

Private Sub Label14_Click()
Unload Me
End Sub

Private Sub optCombine_Click(Index As Integer)
Dim i As Long
Dim sTag As String
   For i = 0 To Me.Controls.Count - 1
      On Error Resume Next
      sTag = Me.Controls(i).Tag
      If (Err.Number = 0) Then
         If (sTag = "Add") Then
            Me.Controls(i).Enabled = optCombine(0).Value
         End If
      End If
   Next i

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



