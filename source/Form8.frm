VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Pistures Opening Window  - Snowman Media  2.0"
   ClientHeight    =   2685
   ClientLeft      =   1800
   ClientTop       =   1230
   ClientWidth     =   4350
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   2685
   ScaleWidth      =   4350
   Begin VB.Image Image1 
      Height          =   2580
      Left            =   0
      Top             =   0
      Width           =   4245
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim w As Integer
Dim h As Integer
Dim a As Integer
Private Sub Form_Load()
Form8.Image1.Stretch = False
a = 1
End Sub
Private Sub form_resize()
If a = 1 Then
w = Form8.Image1.Width
h = Form8.Image1.Height
Form8.Image1.Height = h
Form8.Image1.Width = w
Form8.Image1.Stretch = True
a = 0
End If
If Form8.Image1.Width > Me.ScaleWidth Then
Form8.Image1.Visible = False
Form8.Image1.Width = Form8.Image1.Width * 4 / 5
Form8.Image1.Height = Form8.Image1.Height * 4 / 5
Form8.Image1.Top = Form8.Height / 2 - Form8.Image1.Height / 2
Form8.Image1.Left = Form8.Width / 2 - Form8.Image1.Width / 2
Form8.Image1.Visible = True
End If
Form8.Image1.Visible = False
If Form8.Image1.Height > Me.ScaleHeight Then
Form8.Image1.Height = Form8.Image1.Height * 4 / 5
Form8.Image1.Width = Form8.Image1.Width * 4 / 5
Form8.Image1.Top = Form8.Height / 2 - Form8.Image1.Height / 2
Form8.Image1.Left = Form8.Width / 2 - Form8.Image1.Width / 2
Form8.Image1.Visible = True
End If
If Form8.Image1.Width < Me.ScaleWidth Then
Form8.Image1.Visible = False
Form8.Image1.Width = Form8.Image1.Width * 5 / 4
Form8.Image1.Height = Form8.Image1.Height * 5 / 4
Form8.Image1.Top = Form8.Height / 2 - Form8.Image1.Height / 2
Form8.Image1.Left = Form8.Width / 2 - Form8.Image1.Width / 2
Form8.Image1.Visible = True
End If
If Form8.Image1.Height < Me.ScaleHeight Then
Form8.Image1.Visible = False
Form8.Image1.Height = Form8.Image1.Height * 5 / 4
Form8.Image1.Width = Form8.Image1.Width * 5 / 4
Form8.Image1.Top = Form8.Height / 2 - Form8.Image1.Height / 2
Form8.Image1.Left = Form8.Width / 2 - Form8.Image1.Width / 2
Form8.Image1.Visible = True
End If
End Sub
