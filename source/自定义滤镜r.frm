VERSION 5.00
Begin VB.Form frmCustomFilter 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "自定义滤镜  - Snowman Media Pictures Browser  1.0"
   ClientHeight    =   2700
   ClientLeft      =   6045
   ClientTop       =   3705
   ClientWidth     =   8115
   Icon            =   "自定义滤镜r.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraSep 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   3825
      TabIndex        =   66
      Top             =   2025
      Width           =   4785
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5445
      ScaleHeight     =   255
      ScaleWidth      =   2595
      TabIndex        =   60
      Top             =   1485
      Width           =   2625
      Begin VB.CheckBox chkAuto 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "使用(&D)"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   45
         TabIndex        =   65
         Top             =   0
         Width           =   2625
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3825
      ScaleHeight     =   255
      ScaleWidth      =   1380
      TabIndex        =   59
      Top             =   1485
      Width           =   1410
      Begin VB.Label Label1 
         BackColor       =   &H0080FFFF&
         Caption         =   "自动强化(&A):"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   45
         TabIndex        =   64
         Top             =   45
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3825
      ScaleHeight     =   255
      ScaleWidth      =   1380
      TabIndex        =   58
      Top             =   1035
      Width           =   1410
      Begin VB.Label lblWeight 
         BackColor       =   &H0080FFFF&
         Caption         =   "份量(&W):"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   45
         TabIndex        =   63
         Top             =   45
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3825
      ScaleHeight     =   255
      ScaleWidth      =   1380
      TabIndex        =   57
      Top             =   585
      Width           =   1410
      Begin VB.Label lblBase 
         BackColor       =   &H0080FFFF&
         Caption         =   "基础(&B):"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   45
         TabIndex        =   62
         Top             =   45
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3825
      ScaleHeight     =   255
      ScaleWidth      =   1380
      TabIndex        =   56
      Top             =   135
      Width           =   1410
      Begin VB.Label lblName 
         BackColor       =   &H0080FFFF&
         Caption         =   "标识(&M):"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   45
         TabIndex        =   61
         Top             =   45
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6930
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   54
      Top             =   2205
      Width           =   1095
      Begin VB.Label Label14 
         BackColor       =   &H0000FFFF&
         Caption         =   "取消(&C)"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   45
         TabIndex        =   55
         Top             =   45
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5760
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   52
      Top             =   2205
      Width           =   1095
      Begin VB.Label Label13 
         BackColor       =   &H0000FFFF&
         Caption         =   "确定(&Y)"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   45
         TabIndex        =   53
         Top             =   45
         Width           =   1680
      End
   End
   Begin VB.ComboBox cboName 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   5400
      TabIndex        =   51
      Text            =   "New Filter"
      Top             =   135
      Width           =   2625
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   48
      Left            =   3105
      MaxLength       =   3
      TabIndex        =   50
      Text            =   "0"
      Top             =   2295
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   47
      Left            =   2610
      MaxLength       =   3
      TabIndex        =   49
      Text            =   "0"
      Top             =   2295
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   46
      Left            =   2115
      MaxLength       =   3
      TabIndex        =   48
      Text            =   "0"
      Top             =   2295
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   45
      Left            =   1620
      MaxLength       =   3
      TabIndex        =   47
      Text            =   "0"
      Top             =   2295
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   44
      Left            =   1125
      MaxLength       =   3
      TabIndex        =   46
      Text            =   "0"
      Top             =   2295
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   43
      Left            =   630
      MaxLength       =   3
      TabIndex        =   45
      Text            =   "0"
      Top             =   2295
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   42
      Left            =   135
      MaxLength       =   3
      TabIndex        =   44
      Text            =   "0"
      Top             =   2295
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   41
      Left            =   3105
      MaxLength       =   3
      TabIndex        =   43
      Text            =   "0"
      Top             =   1935
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   40
      Left            =   2610
      MaxLength       =   3
      TabIndex        =   42
      Text            =   "0"
      Top             =   1935
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   39
      Left            =   2115
      MaxLength       =   3
      TabIndex        =   41
      Text            =   "0"
      Top             =   1935
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   38
      Left            =   1620
      MaxLength       =   3
      TabIndex        =   40
      Text            =   "0"
      Top             =   1935
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   37
      Left            =   1125
      MaxLength       =   3
      TabIndex        =   39
      Text            =   "0"
      Top             =   1935
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   36
      Left            =   630
      MaxLength       =   3
      TabIndex        =   38
      Text            =   "0"
      Top             =   1935
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   35
      Left            =   135
      MaxLength       =   3
      TabIndex        =   37
      Text            =   "0"
      Top             =   1935
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   34
      Left            =   3105
      MaxLength       =   3
      TabIndex        =   36
      Text            =   "0"
      Top             =   1575
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   33
      Left            =   2610
      MaxLength       =   3
      TabIndex        =   35
      Text            =   "0"
      Top             =   1575
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   32
      Left            =   2115
      MaxLength       =   3
      TabIndex        =   34
      Text            =   "0"
      Top             =   1575
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   31
      Left            =   1620
      MaxLength       =   3
      TabIndex        =   33
      Text            =   "0"
      Top             =   1575
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   30
      Left            =   1125
      MaxLength       =   3
      TabIndex        =   32
      Text            =   "0"
      Top             =   1575
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   29
      Left            =   630
      MaxLength       =   3
      TabIndex        =   31
      Text            =   "0"
      Top             =   1575
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   28
      Left            =   135
      MaxLength       =   3
      TabIndex        =   30
      Text            =   "0"
      Top             =   1575
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   27
      Left            =   3105
      MaxLength       =   3
      TabIndex        =   29
      Text            =   "0"
      Top             =   1215
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   26
      Left            =   2610
      MaxLength       =   3
      TabIndex        =   28
      Text            =   "0"
      Top             =   1215
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   25
      Left            =   2115
      MaxLength       =   3
      TabIndex        =   27
      Text            =   "0"
      Top             =   1215
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   24
      Left            =   1620
      MaxLength       =   3
      TabIndex        =   26
      Text            =   "1"
      Top             =   1215
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   23
      Left            =   1125
      MaxLength       =   3
      TabIndex        =   25
      Text            =   "0"
      Top             =   1215
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   22
      Left            =   630
      MaxLength       =   3
      TabIndex        =   24
      Text            =   "0"
      Top             =   1215
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   21
      Left            =   135
      MaxLength       =   3
      TabIndex        =   23
      Text            =   "0"
      Top             =   1215
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   20
      Left            =   3105
      MaxLength       =   3
      TabIndex        =   22
      Text            =   "0"
      Top             =   855
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   19
      Left            =   2610
      MaxLength       =   3
      TabIndex        =   21
      Text            =   "0"
      Top             =   855
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   18
      Left            =   2115
      MaxLength       =   3
      TabIndex        =   20
      Text            =   "0"
      Top             =   855
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   17
      Left            =   1620
      MaxLength       =   3
      TabIndex        =   19
      Text            =   "0"
      Top             =   855
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   16
      Left            =   1125
      MaxLength       =   3
      TabIndex        =   18
      Text            =   "0"
      Top             =   855
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   15
      Left            =   630
      MaxLength       =   3
      TabIndex        =   17
      Text            =   "0"
      Top             =   855
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   14
      Left            =   135
      MaxLength       =   3
      TabIndex        =   16
      Text            =   "0"
      Top             =   855
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   13
      Left            =   3105
      MaxLength       =   3
      TabIndex        =   15
      Text            =   "0"
      Top             =   495
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   12
      Left            =   2610
      MaxLength       =   3
      TabIndex        =   14
      Text            =   "0"
      Top             =   495
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   11
      Left            =   2115
      MaxLength       =   3
      TabIndex        =   13
      Text            =   "0"
      Top             =   495
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   10
      Left            =   1620
      MaxLength       =   3
      TabIndex        =   12
      Text            =   "0"
      Top             =   495
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   9
      Left            =   1125
      MaxLength       =   3
      TabIndex        =   11
      Text            =   "0"
      Top             =   495
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   8
      Left            =   630
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "0"
      Top             =   495
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   7
      Left            =   135
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "0"
      Top             =   495
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   6
      Left            =   3105
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "0"
      Top             =   135
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   5
      Left            =   2610
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "0"
      Top             =   135
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   4
      Left            =   2115
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "0"
      Top             =   135
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   3
      Left            =   1620
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "0"
      Top             =   135
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   2
      Left            =   1125
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "0"
      Top             =   135
      Width           =   465
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   1
      Left            =   630
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "0"
      Top             =   135
      Width           =   465
   End
   Begin VB.TextBox txtWeight 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   5400
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "1"
      Top             =   1035
      Width           =   2625
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   0
      Left            =   135
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "0"
      Top             =   135
      Width           =   465
   End
   Begin VB.ComboBox cboBaseOn 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   585
      Width           =   2625
   End
End
Attribute VB_Name = "frmCustomFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cI As New cImageProcessDIB
Private m_bCancel As Boolean

Private Sub pLoadSavedFilter(ByVal lIndex As Long)
Dim cR As New cRegistry
Dim i As Long
Dim vValues As Variant
    cR.ClassKey = HKEY_CURRENT_USER
    cR.SectionKey = "Software\vbAccelerator\vbImageProc\Filter" & lIndex
    If (cR.KeyExists) Then
        cR.ValueKey = "Weight"
        cR.ValueType = REG_DWORD
        txtWeight.Text = cR.Value
        cR.ValueKey = "Values"
        cR.ValueType = REG_BINARY
        vValues = cR.Value
        For i = 0 To 48
            If (vValues(i * 2) <> 0) Then
                txtValue(i) = -1 * vValues(i * 2 + 1)
            Else
                txtValue(i) = vValues(i * 2 + 1)
            End If
        Next i
    End If
    
End Sub
Private Sub pSaveFilter(ByVal sName As String)
Dim lC As Long
Dim cR As New cRegistry
Dim bV() As Byte
Dim iV As Integer
Dim i As Long

    If (cboName.ListIndex = -1) Then
        lC = cboName.ListCount + 1
    Else
        lC = cboName.ItemData(cboName.ListIndex)
    End If
    cR.ClassKey = HKEY_CURRENT_USER
    cR.SectionKey = "Software\vbAccelerator\vbImageProc"
    
    ' Increment the name list:
    cR.ValueKey = "CustomFilters"
    cR.ValueType = REG_DWORD
    cR.Value = lC
    cR.ValueKey = "Filter" & lC
    cR.ValueType = REG_SZ
    cR.Value = sName

    ' Add the specific values:
    cR.SectionKey = cR.SectionKey & "\Filter" & lC
    cR.ValueKey = "Weight"
    cR.ValueType = REG_DWORD
    cR.Value = CLng(txtWeight.Text)
    cR.ValueKey = "Values"
    cR.ValueType = REG_BINARY
    ReDim bV(0 To 98) As Byte
    For i = 0 To 48
        iV = Val(txtValue(i).Text)
        If iV < 0 Then
            bV(i * 2) = 1
        End If
        bV(i * 2 + 1) = Abs(iV)
    Next i
    cR.Value = bV()
    
End Sub

Public Property Get ImageProcess() As cImageProcessDIB
    Set ImageProcess = m_cI
End Property

Public Property Get Cancelled() As Boolean
    Cancelled = m_bCancel
End Property

Private Sub cboBaseOn_Click()
Dim i As Long, j As Long
Dim iIndex As Long
Dim lSize As Long

    If (cboBaseOn.ListIndex <> -1) Then
        If (cboBaseOn.ItemData(cboBaseOn.ListIndex) <> -1) Then
            For i = 0 To 48
                If (i <> 24) Then
                    txtValue(i) = "0"
                    txtValue_Change CInt(i)
                Else
                    txtValue(i) = "1"
                End If
            Next i
            txtWeight = "1"
            If (cboBaseOn.ItemData(cboBaseOn.ListIndex) > -1) Then
                m_cI.FilterType = cboBaseOn.ItemData(cboBaseOn.ListIndex)
                lSize = m_cI.FilterArraySize
                iIndex = 3 - lSize \ 2 + ((3 - lSize \ 2)) * 7
                For i = -lSize \ 2 To lSize \ 2
                    For j = -lSize \ 2 To lSize \ 2
                        txtValue(iIndex) = m_cI.FilterValue(i, j)
                        iIndex = iIndex + 1
                    Next j
                    iIndex = iIndex + 7 - lSize
                Next i
                txtWeight = m_cI.FilterWeight
            End If
        End If
    End If
End Sub

Private Sub cboName_Click()
    If (cboName.ListIndex > -1) Then
        pLoadSavedFilter cboName.ItemData(cboName.ListIndex)
    End If
End Sub

Private Sub chkAuto_Click()
    If (chkAuto.Value = Checked) Then
        pCalculateWeight
    End If
End Sub

Private Sub cmdCancel_Click()
  
End Sub

Private Sub cmdOK_Click()

End Sub

Private Sub Form_Load()
    
    ' Display the options:
    m_bCancel = True
    With cboBaseOn
        .AddItem "<无>"
        .ItemData(.NewIndex) = -1
        .AddItem "<重设>"
        .ItemData(.NewIndex) = -2
        .AddItem "污点"
        .ItemData(.NewIndex) = eBlur
        .AddItem "更多污点"
        .ItemData(.NewIndex) = eBlurMore
        .AddItem "柔化"
        .ItemData(.NewIndex) = eSoften
        .AddItem "深层柔化"
        .ItemData(.NewIndex) = eSoftenMore
        .AddItem "锐化"
        .ItemData(.NewIndex) = eSharpen
        .AddItem "深层锐化"
        .ItemData(.NewIndex) = eSharpenMore
        .ListIndex = 1
        .ListIndex = 0
    End With
    
    ' Load saved values:
    Dim cR As New cRegistry
    Dim lC As Long, lSaved As Long
    
    cR.ClassKey = HKEY_CURRENT_USER
    cR.SectionKey = "Software\vbAccelerator\vbImageProc"
    cR.ValueKey = "CustomFilters"
    cR.Default = 0
    cR.ValueType = REG_DWORD
    lC = cR.Value
    If (lC > 0) Then
        For lSaved = 1 To lC
            cR.ValueKey = "Filter" & lC
            cR.ValueType = REG_SZ
            cboName.AddItem cR.Value
            cboName.ItemData(cboName.NewIndex) = lSaved
        Next lSaved
    End If
    
End Sub

Private Sub Label13_Click()
Dim i As Long, j As Long
Dim iIndex As Long
Dim iMaxI As Long, iMaxJ As Long, iC As Long
Dim iMax As Long
Dim iSIze As Long
Dim bAllZero As Boolean

    m_bCancel = False
    ' Evaluate the size required for this filter:
    bAllZero = True
    For i = 0 To 6
        For j = 0 To 6
            If Val(txtValue(iIndex).Text) <> 0 Then
                bAllZero = False
                iC = Abs(i - 3)
                If (iC > iMaxI) Then iMaxI = iC
                iC = Abs(j - 3)
                If (iC > iMaxJ) Then iMaxJ = iC
            End If
            iIndex = iIndex + 1
        Next j
    Next i
    
    If (bAllZero) Then
        MsgBox "滤镜要求输出值,请输入.", vbInformation
        Exit Sub
    End If
    
    If (iMaxI > iMaxJ) Then iMax = iMaxI Else iMax = iMaxJ
    Debug.Print iMax
    If (iMax < 1) Then
        MsgBox "滤镜最小要 3x3 的大小.", vbInformation
        Exit Sub
    End If
    
    iSIze = iMax * 2 + 1
    ' Store size, weight and filter coefficients:
    m_cI.FilterArraySize = iSIze
    m_cI.FilterWeight = Val(txtWeight.Text)
    
    iIndex = 3 - iSIze \ 2 + ((3 - iSIze \ 2)) * 7
    For i = -iMax To iMax
        For j = -iMax To iMax
            m_cI.FilterValue(i, j) = Val(txtValue(iIndex).Text)
            iIndex = iIndex + 1
        Next j
        iIndex = iIndex + 7 - iSIze
    Next i
        
    pSaveFilter cboName.Text
        
    Unload Me
End Sub

Private Sub Label14_Click()
  m_bCancel = True
    Unload Me
End Sub

Private Sub txtValue_Change(Index As Integer)
    If (IsNumeric(txtValue(Index).Text)) Then
        If (Val(txtValue(Index).Text) = 0) Then
            txtValue(Index).BackColor = vbButtonFace
        Else
            txtValue(Index).BackColor = vbWindowBackground
        End If
    End If
    If (chkAuto.Value = Checked) Then
        pCalculateWeight
    End If
End Sub
Private Sub pCalculateWeight()
Dim i As Long
Dim lWt As Long
    For i = txtValue.LBound To txtValue.UBound
        If IsNumeric(txtValue(i).Text) Then
            lWt = lWt + Val(txtValue(i).Text)
        End If
    Next i
    txtWeight.Text = lWt
End Sub

Private Sub txtValue_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
    Else
        If (KeyAscii <> 8) And (KeyAscii <> Asc("-")) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtWeight_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
    Else
        If (KeyAscii <> 8) And (KeyAscii <> Asc("-")) Then
            KeyAscii = 0
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
