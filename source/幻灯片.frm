VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "幻灯片方式  - Snowman Media Pictures Browser  1.0"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   Icon            =   "幻灯片.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5370
      Left            =   8955
      TabIndex        =   2
      Top             =   1080
      Width           =   1860
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   900
         Top             =   225
      End
      Begin VB.DriveListBox Drive1 
         Height          =   300
         Left            =   0
         TabIndex        =   5
         Top             =   855
         Width           =   1692
      End
      Begin VB.DirListBox Dir1 
         Height          =   1350
         Left            =   0
         TabIndex        =   4
         Top             =   1215
         Width           =   1692
      End
      Begin VB.FileListBox File1 
         Height          =   1890
         Hidden          =   -1  'True
         Left            =   0
         TabIndex        =   3
         Top             =   2745
         Width           =   1572
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   0
      ScaleHeight     =   6255
      ScaleWidth      =   8040
      TabIndex        =   0
      Top             =   0
      Width           =   8040
      Begin VB.Image Image1 
         Height          =   2895
         Left            =   1530
         Top             =   945
         Width           =   4155
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   2220
      TabIndex        =   1
      Top             =   7065
      Width           =   5655
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public filep As String, yy As Boolean, yz As Boolean
Public db As Integer, pic As Boolean, aa As String, yn As Boolean
Dim ss As String
Dim iw, ih, pw, ph, sb As Integer
Sub Time()
mma = Form2.Text1.Text
End Sub









Private Sub Form_Load()
'Form1.ScaleMode = vbPixels
'Picture1.ScaleMode = vbPixels
Me.Show
Picture1.Top = 0
Picture1.Left = 0
fSIZE

Image1.BorderStyle = 0
mmkk = True
Image1.Stretch = True

Frame1.Left = 100000
Form2.Show
End Sub





Sub huang()
On Error GoTo cc
mma = 5
pic = True
If Right(Form1.Dir1, 1) <> "\" Then
aa = Form1.Dir1 + "\"
Else
aa = Form1.Dir1
End If
ss = Dir(aa + "*.*")
Image1.Stretch = False
Image1.Visible = False
If ss > "" Then
Image1.Picture = LoadPicture(aa + ss)
Image1.Top = Me.Height / 2 - Image1.Height / 2
Image1.Left = Me.Width / 2 - Image1.Width / 2
End If

tz

Image1.Stretch = mmkk
Image1.Visible = True
If mma < 1 Then
Timer1.Interval = 4000
Else
Timer1.Interval = mma * 1000
End If
'Else
'pic = False
'Label1.Caption = "自动看图"
'End If
Exit Sub
cc:
Timer1.Interval = 1
End Sub


Private Sub image1_mousedown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then
Dim a, b As Integer
a = x
b = y
Form2.Left = x
Form2.Top = y
Form2.Show
End If
End Sub



Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
Dim a, b As Integer
a = x
b = y
Form2.Left = x
Form2.Top = y
Form2.Show
End If
End Sub

Private Sub Timer1_Timer()
  



On Error GoTo cc
'Form1.Caption = Now
If pic = False Then Exit Sub
ss = Dir
Label2 = Dir
Image1.Stretch = False
Image1.Visible = False
If ss > "" Then

Image1.Picture = LoadPicture(aa + ss)
Image1.Top = Me.Height / 2 - Image1.Height / 2
Image1.Left = Me.Width / 2 - Image1.Width / 2
Else
Image1.Picture = LoadPicture(Dir(aa + "*.*") + ss)
Image1.Top = Me.Height / 2 - Image1.Height / 2
Image1.Left = Me.Width / 2 - Image1.Width / 2
End If
If mma < 1 Then Timer1.Interval = 4000 Else Timer1.Interval = mma * 1000
If mmkk = True Then tz
Image1.Stretch = mmkk
Image1.Visible = True
Exit Sub
cc:
Timer1.Interval = 1
  
  
  

End Sub




Sub tz()
iw = Image1.Width
ih = Image1.Height
pw = Picture1.Width
ph = Picture1.Height
Dim yt As Boolean
If iw > pw Then yt = True
If ih > ph Then yt = True
If yt = False Then Exit Sub
'Image1.Top = 0
'Image1.Left = 0
If iw / pw > ih / ph Then
Image1.Width = pw
Image1.Height = ih * (pw / iw)
Else
Image1.Width = iw * (ph / ih)
Image1.Height = ph
End If
End Sub
Sub fSIZE()
Picture1.Width = Form1.Width
Picture1.Height = Form1.Height
End Sub
