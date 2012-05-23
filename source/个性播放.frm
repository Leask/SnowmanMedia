VERSION 5.00
Begin VB.Form Form100 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Snowman Media ilxz 4 Playing..."
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1140
   ForeColor       =   &H00000000&
   Icon            =   "个性播放.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   1095
   ScaleWidth      =   1140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   1080
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   1485
      Width           =   5325
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   15
         Left            =   0
         OLEDropMode     =   1  'Manual
         Picture         =   "个性播放.frx":2CFA
         Top             =   0
         Width           =   15
      End
   End
End
Attribute VB_Name = "Form100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_DblClick()
Form1.a003010.Checked = True
Form1.a003010_Click
End Sub

Private Sub Form_GotFocus()
Form1.WindowState = 0
Form1.Ly.CenterForm Form1

Form1.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Form1.WindowState = 0
Form1.Ly.CenterForm Form1

Form1.Show
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu Form1.b002, 0, x, y
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim ThisFile As Variant
Form1.LF1.Clear
Form1.LF2.Clear

For Each ThisFile In Data.Files
Form1.LF1.AddItem ThisFile
       Form1.AddFile Form1.LF1.List(Form1.LF1.ListCount - 1)

Next
Form1.Pid = 0
Form1.LF1.ListIndex = Form1.Pid
Form1.LF1_DblClick


End Sub

Private Sub Form_Resize()
Image1.Picture = LoadPicture(App.Path + "\SmM_Pictures\RightDown.gif")
Frame1.Left = Form100.Width - Frame1.Width - 20
Frame1.Top = Form100.Height - Frame1.Height - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Form100 = Nothing
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Form1.WindowState = 0
Form1.Ly.CenterForm Form1

Form1.Show

End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu Form1.b002, 0, Frame1.Left + x, Frame1.Top + y

End Sub

Private Sub Frame1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim ThisFile As Variant
Form1.LF1.Clear
Form1.LF2.Clear

For Each ThisFile In Data.Files
Form1.LF1.AddItem ThisFile
Form1.AddFile Form1.LF1.List(Form1.LF1.ListCount - 1)

Next
Form1.Pid = 0
Form1.LF1.ListIndex = Form1.Pid
Form1.LF1_DblClick

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Form1.WindowState = 0
 Form1.Ly.CenterForm Form1

Form1.Show

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu Form1.b002, 0, Frame1.Left + x, Frame1.Top + y

End Sub

Private Sub Image1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim ThisFile As Variant
Form1.LF1.Clear
Form1.LF2.Clear

For Each ThisFile In Data.Files
Form1.LF1.AddItem ThisFile
Form1.AddFile Form1.LF1.List(Form1.LF1.ListCount - 1)

Next
Form1.Pid = 0
Form1.LF1.ListIndex = Form1.Pid
Form1.LF1_DblClick

End Sub

