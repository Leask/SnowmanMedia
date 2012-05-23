VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "Snowman Media Scaner  1.0"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   Icon            =   "É¨Ãè.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   7200
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   16380
      TabIndex        =   1
      Top             =   0
      Width           =   16383
      Begin VB.PictureBox Picture3 
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
         Left            =   3510
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   6
         Top             =   45
         Width           =   1365
         Begin VB.Label Label1 
            BackColor       =   &H0000FFFF&
            Caption         =   "ÍË³ö(&X)"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   45
            TabIndex        =   7
            Top             =   45
            Width           =   1635
         End
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
         Left            =   180
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   4
         Top             =   45
         Width           =   1365
         Begin VB.Label Label13 
            BackColor       =   &H0000FFFF&
            Caption         =   "Ñ¡ÔñÆ÷²Ä(&T)"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   45
            TabIndex        =   5
            Top             =   45
            Width           =   1230
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
         Left            =   1755
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   2
         Top             =   45
         Width           =   1365
         Begin VB.Label Label14 
            BackColor       =   &H0000FFFF&
            Caption         =   "É¨Ãè(&S)"
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   45
            TabIndex        =   3
            Top             =   45
            Width           =   1635
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   6585
      Left            =   0
      Picture         =   "É¨Ãè.frx":1582
      ScaleHeight     =   6555
      ScaleWidth      =   8985
      TabIndex        =   0
      Top             =   360
      Width           =   9015
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label13_Click()
 r = TWAIN_SelectImageSource(Me.hWnd)
End Sub

Private Sub Label14_Click()
r = TWAIN_AcquireToClipboard(Me.hWnd, t%)
  Picture1.Picture = Clipboard.GetData(vbCFDIB)
End Sub
Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label14.BackColor = &HFF0000
Picture10.BackColor = &HFF0000
Label14.ForeColor = &HFFFF&
End Sub
Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label13.BackColor = &HFF0000
Picture9.BackColor = &HFF0000
Label13.ForeColor = &HFFFF&
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.BackColor = &HFF0000
Picture3.BackColor = &HFF0000
Label1.ForeColor = &HFFFF&
End Sub

Private Sub picture2_mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture9.BackColor = &HFFFF&
Picture3.BackColor = &HFFFF&
Picture10.BackColor = &HFFFF&
Label13.ForeColor = &HFF0000
Label13.BackColor = &HFFFF&
Label14.ForeColor = &HFF0000
Label14.BackColor = &HFFFF&
Label1.ForeColor = &HFF0000
Label1.BackColor = &HFFFF&

End Sub
Private Sub picture1_mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture9.BackColor = &HFFFF&
Picture3.BackColor = &HFFFF&
Picture10.BackColor = &HFFFF&
Label13.ForeColor = &HFF0000
Label13.BackColor = &HFFFF&
Label14.ForeColor = &HFF0000
Label14.BackColor = &HFFFF&
Label1.ForeColor = &HFF0000
Label1.BackColor = &HFFFF&

End Sub


