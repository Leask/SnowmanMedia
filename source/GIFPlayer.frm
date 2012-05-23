VERSION 5.00
Object = "{220F55E8-7AAE-11D3-9D68-F74ED5721646}#18.0#0"; "TANI.OCX"
Object = "{F85233B2-B49F-11D2-BD06-EB39B7A2BD6C}#2.0#0"; "EASYSCROLL.OCX"
Begin VB.Form FrmAniGif 
   BackColor       =   &H00000000&
   Caption         =   "Gif  £­ Snowman Media Pictures Browser  1.0"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6060
   ForeColor       =   &H00000000&
   Icon            =   "GIFPlayer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   8565
      Left            =   0
      ScaleHeight     =   8535
      ScaleWidth      =   2100
      TabIndex        =   1
      Top             =   0
      Width           =   2130
      Begin EasyScroll_ActiveX_Control.EasyScroll EasyScroll2 
         Height          =   420
         Left            =   45
         Top             =   45
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   741
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   8505
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2085
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            MouseIcon       =   "GIFPlayer.frx":1582
            MousePointer    =   99  'Custom
            ScaleHeight     =   255
            ScaleWidth      =   1695
            TabIndex        =   7
            Top             =   8505
            Width           =   1725
            Begin VB.Label Label5 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "ÍË³ö(&X)"
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   45
               MouseIcon       =   "GIFPlayer.frx":16D4
               TabIndex        =   8
               Top             =   45
               Width           =   1950
            End
         End
         Begin VB.DriveListBox Drive1 
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   90
            MouseIcon       =   "GIFPlayer.frx":1826
            TabIndex        =   6
            Top             =   1125
            Width           =   1725
         End
         Begin VB.DirListBox Dir1 
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   2820
            Left            =   90
            MouseIcon       =   "GIFPlayer.frx":1978
            TabIndex        =   5
            Top             =   1665
            Width           =   1725
         End
         Begin VB.FileListBox File1 
            BackColor       =   &H0080FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   3150
            Left            =   90
            MouseIcon       =   "GIFPlayer.frx":1ACA
            Pattern         =   "*.gif"
            TabIndex        =   4
            Top             =   4725
            Width           =   1725
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   90
            TabIndex        =   3
            Top             =   450
            Width           =   1725
         End
      End
   End
   Begin TAni.TMaxAni TMaxAni1 
      Height          =   510
      Left            =   2295
      TabIndex        =   0
      Top             =   180
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   900
   End
End
Attribute VB_Name = "FrmAniGif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FileSelect$
Private Sub Dir1_Change()
On Error Resume Next
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Text1.Text = File1.FileName
If Len(File1.Path) > 3 Then
    FileSelect$ = File1.Path & "\" & File1.FileName
Else
    FileSelect$ = File1.Path & File1.FileName
End If
Me.Caption = FileSelect$ + "  - Snowman Media Pictures Browser  1.0"
End Sub
Private Sub Dir1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dir1.BackColor = &HFF0000
Dir1.ForeColor = &HFFFF&
End Sub
Private Sub File1_DblClick()
On Error Resume Next
TMaxAni1.FileName = FileSelect$
TMaxAni1.ShowGif
ResizeForm
TMaxAni1.Left = 2295
End Sub
Sub ResizeForm()

    Me.Height = TMaxAni1.Height + 800
    Me.Width = TMaxAni1.Width + 2600
    If Me.Height < 4980 Then Me.Height = 4980
    If Me.Width < 6180 Then Me.Width = 6180


End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
File1.BackColor = &HFF0000
File1.ForeColor = &HFFFF&
End Sub

Private Sub Form_Load()
TMaxAni1.Left = 10000
End Sub
Private Sub Form_Resize()
Picture1.Height = Me.Height - 150
Frame1.Height = Picture1.Height + 4000

End Sub

Private Sub Label5_Click()
Unload Me
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5.BackColor = &HFF0000
Picture4.BackColor = &HFF0000
Label5.ForeColor = &HFFFF&
End Sub
Private Sub frame1_mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
File1.BackColor = &HFFFF&
File1.ForeColor = &HFF0000
Label5.BackColor = &HFFFF&
Label5.ForeColor = &HFF0000
Picture4.BackColor = &HFFFF&
Dir1.BackColor = &HFFFF&
Dir1.ForeColor = &HFF0000
Text1.BackColor = &HFFFF&
Text1.ForeColor = &HFF0000
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Text1.BackColor = &HFFFF&
Text1.ForeColor = &HFF0000
End Sub
