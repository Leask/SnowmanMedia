VERSION 5.00
Begin VB.Form Form103 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Now Starting  Snowman Media  2.0"
   ClientHeight    =   3735
   ClientLeft      =   6330
   ClientTop       =   660
   ClientWidth     =   6060
   Icon            =   "frmfade.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3600
      Top             =   4815
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   200
      LargeChange     =   5
      Left            =   2115
      Max             =   100
      SmallChange     =   5
      TabIndex        =   2
      Top             =   5715
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3735
      ScaleMode       =   0  'User
      ScaleWidth      =   6060
      TabIndex        =   1
      Top             =   0
      Width           =   6060
   End
   Begin VB.PictureBox Picture2 
      Height          =   3105
      Left            =   2790
      ScaleHeight     =   3045
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   1485
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   990
      Picture         =   "frmfade.frx":1582
      ScaleHeight     =   2265
      ScaleWidth      =   4995
      TabIndex        =   3
      Top             =   630
      Visible         =   0   'False
      Width           =   5025
   End
End
Attribute VB_Name = "Form103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Dim st As Integer
Dim patfade As PatternFade
Private Sub Explode(Newform As Form, Increment As Integer)
Dim Size As RECT                                                                                                ' setup form as rect type
GetWindowRect Newform.hWnd, Size
Dim FormWidth, FormHeight As Integer                                                                 ' establish dimension variables
FormWidth = (Size.Right - Size.Left)
FormHeight = (Size.Bottom - Size.Top)
Dim TempDC
TempDC = GetDC(ByVal 0&)                                                                                 ' obtain memory dc for resizing
Dim Count, LeftPoint, TopPoint, nWidth, nHeight As Integer                                      ' establish resizing variables
For Count = 1 To Increment                                                                               ' loop to new sizes
    nWidth = FormWidth * (Count / Increment)
    nHeight = FormHeight * (Count / Increment)
    LeftPoint = Size.Left + (FormWidth - nWidth) / 2
    TopPoint = Size.Top + (FormHeight - nHeight) / 2
    Rectangle TempDC, LeftPoint, TopPoint, LeftPoint + nWidth, TopPoint + nHeight     ' draw rectangles to build form
Next Count
DeleteDC (TempDC)                                                                                           ' release  memory resource
End Sub
Private Sub Form_Load()
     Form103.Picture3.Width = 6060
     Form103.Picture3.Height = 3735
     Form103.Picture3.Top = 0
     Form103.Picture3.Left = 0
   Move (Screen.Width - Form103.Width) \ 2, (Screen.Height - Form103.Height) \ 2
    Set patfade = New PatternFade
    Set patfade.pic1 = Picture1
    Set patfade.pic2 = Picture2
    Set patfade.pic3 = Picture3
    patfade.Setup
    HScroll1.Value = 50
End Sub
Private Sub form_resize()
  Explode Me, 1000
   'patfade.FadeOut HScroll1.Value
 patfade.FadeIn HScroll1.Value
End Sub
Private Sub Timer1_Timer()
st = st + 1
If st = 5 Then
  Form102.Show
  Unload Me
End If
End Sub
