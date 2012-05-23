VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   0  'None
   Caption         =   "Form9"
   ClientHeight    =   495
   ClientLeft      =   3870
   ClientTop       =   1560
   ClientWidth     =   1980
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   1980
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   1620
      Left            =   0
      Picture         =   "Form9.frx":0000
      ToolTipText     =   "你可以把我拖到任何地方，单击显示""雪人""，或双击取消我的最小化  －－Snowman Media"
      Top             =   -135
      Width           =   1980
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim intpos As Integer
Const hwnd_top = 0
Dim ingreturnvalue As Long
Const HWND_TOPMOST = -1
Const SWP_SHOWWINDOW = &H40
Private Sub Form_Load()
Me.Height = Form9.Image1.Height
Me.Width = Form9.Image1.Width
ingreturnvalue = SetWindowPos(Me.hwnd, HWND_TOPMOST, Val(10), Val(10), Val(10), Val(10), SWP_SHOWWINDOW)
Form9.Width = 1980
Form9.Height = 495
Form9.Caption = "Snowman Media"
Form9.Top = 6000
Form9.Left = 8888
End Sub
Private Sub image1_DblClick()
Form1.Visible = True
Form1.Show
End Sub

Private Sub image1_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub image1_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form9.Left = Form9.Left + (c - a)
Form9.Top = Form9.Top + (d - b)
End Sub
Private Sub form_mousedown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button = 1 Then
a = X
b = Y
End If
End Sub
Private Sub form_mousemove(Button As Integer, shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub

c = X
d = Y
Form9.Left = Form9.Left + (c - a)
Form9.Top = Form9.Top + (d - b)
End Sub

