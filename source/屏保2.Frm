VERSION 5.00
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "SmM_Tools.ocx"
Begin VB.Form form100 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Snowman Media Snow"
   ClientHeight    =   4455
   ClientLeft      =   -795
   ClientTop       =   0
   ClientWidth     =   4500
   ClipControls    =   0   'False
   Icon            =   "屏保2.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   990
      TabIndex        =   0
      Top             =   2250
      Width           =   5325
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   1050
         Left            =   0
         Picture         =   "屏保2.frx":2CFA
         Top             =   0
         Width           =   5325
      End
   End
   Begin API控制大全.LyfTools LyfTools1 
      Left            =   2655
      Top             =   1395
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "form100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Long, b As Long, c As Long, d As Long, e As Long
Dim j As Long
Dim ls As Long
Dim i As Long
Dim oldx As Long
Dim oldy As Long
Dim Coloury As Long
Dim Snow(1000, 2), Amounty As Long
Sub SnowA()
Call CheckExist(Me)
form100.LyfTools1.EnableCursor False
e = 1
form100.Show
DoEvents
Randomize: Amounty = 325
For j = 1 To Amounty
Snow(j, 0) = Int(Rnd * form100.Width)
Snow(j, 1) = Int(Rnd * form100.Height)
Snow(j, 2) = 10 + (Rnd * 20)
Next j
Do While Not (DoEvents = 0)
For ls = 1 To 10
For i = 1 To Amounty
oldx = Snow(i, 0): oldy = Snow(i, 1): Snow(i, 1) = Snow(i, 1) + Snow(i, 2)
If Snow(i, 1) > form100.Height Then Snow(i, 1) = 0: Snow(i, 2) = 5 + (Rnd * 30): Snow(i, 0) = Int(Rnd * form100.Width): oldx = 0: oldy = 0
Coloury = 8 * (Snow(i, 2) - 10): Coloury = 60 + Coloury: PSet (oldx, oldy), QBColor(0): PSet (Snow(i, 0), Snow(i, 1)), RGB(Coloury, Coloury, Coloury)
Next i
Next ls
Frame1.Refresh
'Label2.Refresh
'Label3.Refresh
Loop
End

End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End
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

Private Sub Form_Resize()
Frame1.Left = form100.Width - Frame1.Width - 20
Frame1.Top = form100.Height - Frame1.Height - 50
SnowA
End Sub
