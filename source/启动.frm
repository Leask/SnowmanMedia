VERSION 5.00
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "LYFTOOLS.OCX"
Begin VB.Form Form103 
   Appearance      =   0  'Flat
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Snowman Media  3.0"
   ClientHeight    =   3690
   ClientLeft      =   6330
   ClientTop       =   660
   ClientWidth     =   6015
   Icon            =   "启动.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.HScrollBar HScroll1 
      Height          =   200
      LargeChange     =   5
      Left            =   3465
      Max             =   100
      SmallChange     =   5
      TabIndex        =   2
      Top             =   6705
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   3690
      Left            =   0
      ScaleHeight     =   3690
      ScaleMode       =   0  'User
      ScaleWidth      =   6015
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   4725
         Top             =   3195
      End
      Begin API控制大全.LyfTools LyfTools1 
         Left            =   5355
         Top             =   3060
         _ExtentX        =   847
         _ExtentY        =   847
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   3690
      ScaleHeight     =   360
      ScaleWidth      =   525
      TabIndex        =   0
      Top             =   5580
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   3675
      Left            =   2025
      Picture         =   "启动.frx":1582
      ScaleHeight     =   3675
      ScaleWidth      =   6000
      TabIndex        =   3
      Top             =   5490
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   """"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   810
      TabIndex        =   4
      Top             =   6345
      Visible         =   0   'False
      Width           =   1770
   End
End
Attribute VB_Name = "Form103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UlM As Boolean
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Dim st As String
Dim patfade As PatternFade
Private Sub Explode(Newform As Form, Increment As Integer)
On Error Resume Next
Dim Size As RECT                                                                                                ' setup form as rect type
GetWindowRect Newform.hwnd, Size
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

On Error Resume Next
Me.LyfTools1.MakeTop Me, True
Timer1.Enabled = True

    Set patfade = New PatternFade
    Set patfade.pic1 = Picture1
    Set patfade.pic2 = Picture2
    Set patfade.pic3 = Picture3
    HScroll1.Value = 50
   patfade.Setup
        LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Filename", "NoFile"

  End Sub

Private Sub Form_Resize()
On Error Resume Next

  Explode Me, 1000
   patfade.FadeIn HScroll1.Value
 If Len(Command) > 0 Then
st = sReplace(Command, Label1.Caption, "")
 st = sReplace(st, Label1.Caption, "")
     LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Filename", st
 Else: LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Filename", "NoFile"

 End If

If LyfTools1.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Starting") <> True Then Shell (App.Path + "\SmM_Mp.exe")
LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Starting", False
UlM = True
 'txtstrRem.Text = Label1.Caption
' Dim s As New clsStrings
' Form102.MediaPlayer1.Filename = s.RemoveString(txtstrRem, txtstrToRem)
 'Form102.Show
 
End Sub
Function sReplace(SearchLine As String, SearchFor As String, ReplaceWith As String)
On Error Resume Next

Dim vSearchLine As String, found As Integer

found = InStr(SearchLine, SearchFor): vSearchLine = SearchLine
If found <> 0 Then
vSearchLine = ""
If found > 1 Then vSearchLine = Left(SearchLine, found - 1)
vSearchLine = vSearchLine + ReplaceWith
If found + Len(SearchFor) - 1 < Len(SearchLine) Then _
vSearchLine = vSearchLine + Right$(SearchLine, Len(SearchLine) - found - Len(SearchFor) + 1)
End If
sReplace = vSearchLine

End Function


Private Sub Timer1_Timer()
On Error Resume Next
If UlM = True Then
Unload Me
End
End If
End Sub
