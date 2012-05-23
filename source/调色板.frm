VERSION 5.00
Begin VB.Form frmPalette 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "调色板  － Snowman Media P.B.  1.0"
   ClientHeight    =   5820
   ClientLeft      =   6525
   ClientTop       =   3930
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "调色板.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox cboPalette 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Height          =   300
      Left            =   135
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4950
      Width           =   4380
   End
   Begin VB.PictureBox picPalette 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      Height          =   4380
      Left            =   135
      ScaleHeight     =   4350
      ScaleWidth      =   4350
      TabIndex        =   7
      Top             =   90
      Width           =   4380
   End
   Begin VB.Frame fraSep 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   15
      Left            =   -495
      TabIndex        =   6
      Top             =   5400
      Width           =   7260
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2250
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   4
      Top             =   5490
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
         TabIndex        =   5
         Top             =   45
         Width           =   1680
      End
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3420
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   2
      Top             =   5490
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
         TabIndex        =   3
         Top             =   45
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   135
      ScaleHeight     =   255
      ScaleWidth      =   1560
      TabIndex        =   0
      Top             =   4590
      Width           =   1590
      Begin VB.Label lblPalette 
         BackColor       =   &H0080FFFF&
         Caption         =   "调色板(&P):"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   45
         TabIndex        =   1
         Top             =   45
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sSelected As String
Private m_bCancel As Boolean
Private m_sName() As String
Private m_sFIle() As String
Private m_iCount As Long
Private m_cPal As New cPalette

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Public Property Get FileName() As String
   FileName = m_sSelected
End Property
Public Property Get Cancelled() As Boolean
   Cancelled = m_bCancel
End Property

Private Sub LoadPalette()
On Error GoTo LoadPaletteError
   picPalette.Cls
   If m_cPal.LoadFromFile(m_sSelected) Then
      RenderPalette
   End If
   Exit Sub
LoadPaletteError:
   MsgBox "Snowman Media Pictures Browser  1.0 在加载图片时发生错误:" & Err.Description
   Exit Sub
End Sub

Private Sub RenderPalette()
Dim iPal As Long
Dim x As Long, y As Long
Dim lHDC As Long
Dim hBR As Long
Dim tR As RECT
   lHDC = picPalette.hdc
   x = 3: y = 3
   For iPal = 1 To m_cPal.Count
      tR.Left = x: tR.Right = tR.Left + 14
      tR.Top = y: tR.Bottom = tR.Top + 14
      hBR = CreateSolidBrush(RGB(m_cPal.Red(iPal), m_cPal.Green(iPal), m_cPal.Blue(iPal)))
      FillRect lHDC, tR, hBR
      DeleteObject hBR
      x = x + 18
      If (x > 290) Then
         x = 3
         y = y + 18
      End If
   Next iPal
   picPalette.Refresh
End Sub

Private Sub SortPalette()
   ' Todo...
   RenderPalette
End Sub

Private Sub cboPalette_Click()
   If (cboPalette.ListIndex > -1) Then
      m_sSelected = m_sFIle(cboPalette.ItemData(cboPalette.ListIndex))
      LoadPalette
   End If
End Sub

Private Sub cmdCancel_Click()
 
End Sub

Private Sub cmdOK_Click()
 
End Sub

Private Sub Form_Load()
Dim i As Long

   m_bCancel = True
   
   ReDim m_sName(1 To 1) As String
   ReDim m_sFIle(1 To 1) As String
   m_iCount = 1
   m_sName(1) = "Microsoft(R) Internet Explorer  256色调色板"
   m_sFIle(1) = App.Path & "\216ie.pal"

   For i = 1 To m_iCount
      cboPalette.AddItem m_sName(i)
      cboPalette.ItemData(cboPalette.NewIndex) = i
   Next i
   cboPalette.ListIndex = 0
   
End Sub

Private Sub Label13_Click()
  If (cboPalette.ListIndex < 0) Then
      MsgBox "请先选择一个调色板.", vbInformation
   Else
      m_bCancel = False
      Unload Me
   End If
End Sub

Private Sub Label14_Click()
  Unload Me
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


