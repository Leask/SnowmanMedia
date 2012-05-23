VERSION 5.00
Begin VB.Form form100 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Snow - Peter Chamberlin"
   ClientHeight    =   4365
   ClientLeft      =   3705
   ClientTop       =   3300
   ClientWidth     =   7080
   ClipControls    =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4365
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
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
         Caption         =   "6.2.2001"
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
         Left            =   6075
         TabIndex        =   3
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Snowman Media  2.0"
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
         Left            =   3375
         TabIndex        =   2
         Top             =   45
         Width           =   3360
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(C)Copyright 2000-2001  H2ont   www.h2ont.com h2ont@china.com"
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
         Left            =   945
         TabIndex        =   1
         Top             =   450
         Width           =   5820
      End
   End
End
Attribute VB_Name = "form100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim snow(1000, 2), Amounty As Integer
Private Sub Form_Load()
form100.Show
form100.Frame1.Left = form100.Width - 6830
form100.Frame1.Top = form100.Height - 1050
DoEvents
Randomize: Amounty = 325
For J = 1 To Amounty
snow(J, 0) = Int(Rnd * form100.Width)
snow(J, 1) = Int(Rnd * form100.Height)
snow(J, 2) = 10 + (Rnd * 20)
Next J
Do While Not (DoEvents = 0)
For LS = 1 To 10
For i = 1 To Amounty
OldX = snow(i, 0): OldY = snow(i, 1): snow(i, 1) = snow(i, 1) + snow(i, 2)
If snow(i, 1) > form100.Height Then snow(i, 1) = 0: snow(i, 2) = 5 + (Rnd * 30): snow(i, 0) = Int(Rnd * form100.Width): OldX = 0: OldY = 0
Coloury = 8 * (snow(i, 2) - 10): Coloury = 60 + Coloury: PSet (OldX, OldY), QBColor(0): PSet (snow(i, 0), snow(i, 1)), RGB(Coloury, Coloury, Coloury)
Next i
Next LS
Label1.Refresh
Label2.Refresh
Label3.Refresh
Loop
End
End Sub
'Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 Then
'Unload Me
'End If
'End Sub


