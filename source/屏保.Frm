VERSION 5.00
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "LYFTOOLS.OCX"
Begin VB.Form form100 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Snowman Media Snow"
   ClientHeight    =   7215
   ClientLeft      =   3705
   ClientTop       =   3300
   ClientWidth     =   8130
   ClipControls    =   0   'False
   Icon            =   "屏保.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7215
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin API控制大全.LyfTools LyfTools1 
      Left            =   4320
      Top             =   2430
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   180
      TabIndex        =   0
      Top             =   495
      Width           =   6810
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "10.30.2001"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   5850
         TabIndex        =   3
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Snowman Media ilxz 3.5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   2880
         TabIndex        =   2
         Top             =   45
         Width           =   3870
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (C) 2000-2001 H2ont Leask  www.h2ont.com leask@21cn.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   450
         TabIndex        =   1
         Top             =   450
         Width           =   6330
      End
   End
End
Attribute VB_Name = "form100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a, b, c, d, e As Integer
Dim j As Integer
Dim ls As Integer
Dim i As Integer
Dim oldx As Integer
Dim oldy As Integer
Dim Coloury As Integer
Dim snow(1000, 2), Amounty As Integer
Private Sub Form_Load()
Call CheckExist(Me)
form100.LyfTools1.EnableCursor False
e = 1
form100.Show
form100.Frame1.Left = form100.Width - 6900
form100.Frame1.Top = form100.Height - 1050
DoEvents
Randomize: Amounty = 325
For j = 1 To Amounty
snow(j, 0) = Int(Rnd * form100.Width)
snow(j, 1) = Int(Rnd * form100.Height)
snow(j, 2) = 10 + (Rnd * 20)
Next j
Do While Not (DoEvents = 0)
For ls = 1 To 10
For i = 1 To Amounty
oldx = snow(i, 0): oldy = snow(i, 1): snow(i, 1) = snow(i, 1) + snow(i, 2)
If snow(i, 1) > form100.Height Then snow(i, 1) = 0: snow(i, 2) = 5 + (Rnd * 30): snow(i, 0) = Int(Rnd * form100.Width): oldx = 0: oldy = 0
Coloury = 8 * (snow(i, 2) - 10): Coloury = 60 + Coloury: PSet (oldx, oldy), QBColor(0): PSet (snow(i, 0), snow(i, 1)), RGB(Coloury, Coloury, Coloury)
Next i
Next ls
Label1.Refresh
Label2.Refresh
Label3.Refresh
Loop
End
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If e = 1 Then
a = x
b = y
e = 2
End If
If x <> a Or y <> b Then
End
End If
End Sub
Private Sub Form_Click()
End
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
End
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
End
End Sub
Public Sub CheckExist(fm As Form)
    If App.PrevInstance Then
        End
    End If
End Sub
