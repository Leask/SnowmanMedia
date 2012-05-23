VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "SmM. P.B.  1.0"
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1740
   Icon            =   "菜单.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6465
   ScaleWidth      =   1740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "SmM. P.B.  1.0"
      ForeColor       =   &H80000008&
      Height          =   6450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1725
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         FillColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   90
         MouseIcon       =   "菜单.frx":1582
         ScaleHeight     =   255
         ScaleWidth      =   1515
         TabIndex        =   13
         Top             =   1080
         Width           =   1545
         Begin VB.Label Label6 
            BackColor       =   &H0000FFFF&
            Caption         =   "设置更新时间(&R)"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   45
            MouseIcon       =   "菜单.frx":16D4
            TabIndex        =   3
            Top             =   45
            Width           =   1500
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   990
         TabIndex        =   4
         Text            =   "5"
         Top             =   1440
         Width           =   645
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   90
         MouseIcon       =   "菜单.frx":1826
         ScaleHeight     =   255
         ScaleWidth      =   1515
         TabIndex        =   12
         Top             =   1800
         Width           =   1545
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            Caption         =   "退出(&X)"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   45
            MouseIcon       =   "菜单.frx":1978
            TabIndex        =   5
            Top             =   45
            Width           =   1950
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   675
         Top             =   5355
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   90
         MouseIcon       =   "菜单.frx":1ACA
         ScaleHeight     =   255
         ScaleWidth      =   1515
         TabIndex        =   11
         Top             =   2520
         Width           =   1545
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            Caption         =   "更改目录(&R)"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   45
            TabIndex        =   6
            Top             =   45
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   90
         MouseIcon       =   "菜单.frx":1C1C
         ScaleHeight     =   255
         ScaleWidth      =   1515
         TabIndex        =   10
         Top             =   630
         Width           =   1545
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            Caption         =   "停止(&C)"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   45
            MouseIcon       =   "菜单.frx":1D6E
            TabIndex        =   2
            Top             =   45
            Width           =   1950
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   90
         MouseIcon       =   "菜单.frx":1EC0
         ScaleHeight     =   255
         ScaleWidth      =   1515
         TabIndex        =   9
         Top             =   270
         Width           =   1545
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            Caption         =   "开始(&S)"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   45
            MouseIcon       =   "菜单.frx":2012
            TabIndex        =   1
            Top             =   45
            Width           =   1950
         End
      End
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   90
         MouseIcon       =   "菜单.frx":2164
         TabIndex        =   7
         Top             =   2925
         Width           =   1560
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   3030
         Left            =   90
         MouseIcon       =   "菜单.frx":22B6
         TabIndex        =   8
         Top             =   3330
         Width           =   1560
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         Height          =   60
         Left            =   90
         Top             =   2250
         Width           =   1545
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FFFF&
         Caption         =   " 妙/幅(&S)"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   90
         MouseIcon       =   "菜单.frx":2408
         TabIndex        =   14
         Top             =   1440
         Width           =   1545
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer
Dim e, b, c, d As Integer



Private Sub Form_Load()
a = 1
End Sub

Private Sub frame2_mousedown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
e = x
b = y
End If
End Sub
Private Sub frame2_mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.BackColor = &HFFFF&
Label4.BackColor = &HFFFF&
Label2.BackColor = &HFFFF&
Label6.BackColor = &HFFFF&
Label5.BackColor = &HFFFF&
Label3.ForeColor = &HFF0000
Label4.ForeColor = &HFF0000
Label2.ForeColor = &HFF0000
Label6.ForeColor = &HFF0000
Label5.ForeColor = &HFF0000
Picture1.BackColor = &HFFFF&
Picture2.BackColor = &HFFFF&
Picture3.BackColor = &HFFFF&
Picture5.BackColor = &HFFFF&
Picture4.BackColor = &HFFFF&
Label7.BackColor = &HFFFF&
Text1.BackColor = &HFFFF&
Label7.ForeColor = &HFF0000
Text1.ForeColor = &HFF0000



If Button <> 1 Then Exit Sub
c = x
d = y
Me.Left = Me.Left + (c - e)
Me.Top = Me.Top + (d - b)


End Sub







Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label7.BackColor = &HFF0000
Text1.BackColor = &HFF0000
Label7.ForeColor = &HFFFF&
Text1.ForeColor = &HFFFF&
End Sub

Private Sub Label2_Click()

Form1.Timer1.Enabled = False

End Sub

Private Sub Label3_Click()
Form1.Dir1.Path = Dir1.Path
Call Form1.huang
End Sub

Private Sub Label4_Click()

Form1.Timer1.Enabled = True


End Sub
Private Sub Drive1_Change()
On Error Resume Next
Dir1 = Drive1.Drive
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label3.BackColor = &HFF0000
Picture2.BackColor = &HFF0000
Label3.ForeColor = &HFFFF&
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label4.BackColor = &HFF0000
Picture1.BackColor = &HFF0000
Label4.ForeColor = &HFFFF&
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.BackColor = &HFF0000
Picture3.BackColor = &HFF0000
Label2.ForeColor = &HFFFF&
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label6.BackColor = &HFF0000
Picture5.BackColor = &HFF0000
Label6.ForeColor = &HFFFF&
End Sub


Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5.BackColor = &HFF0000
Picture4.BackColor = &HFF0000
Label5.ForeColor = &HFFFF&
End Sub





Private Sub Label5_Click()
Unload Form1
Unload Me
End Sub

Private Sub Label6_Click()
Call Form1.Time
End Sub

Private Sub Timer1_Timer()
a = a + 1

If a = 2 Then
Call Form1.huang
Timer1.Enabled = False
End If
End Sub
