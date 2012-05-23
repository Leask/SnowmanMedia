VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "关于 Snowman Media ilxz 4"
   ClientHeight    =   3225
   ClientLeft      =   3615
   ClientTop       =   5550
   ClientWidth     =   6000
   ForeColor       =   &H00000000&
   Icon            =   "关于.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3225
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4500
      Top             =   5400
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   3225
      Left            =   0
      ScaleHeight     =   215
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public Ext As Boolean
Const DT_BOTTOM As Long = &H8
Const DT_CALCRECT As Long = &H400
Const DT_CENTER As Long = &H1
Const DT_EXPANDTABS As Long = &H40
Const DT_EXTERNALLEADING As Long = &H200
Const DT_LEFT As Long = &H0
Const DT_NOCLIP As Long = &H100
Const DT_NOPREFIX As Long = &H800
Const DT_RIGHT As Long = &H2
Const DT_SINGLELINE As Long = &H20
Const DT_TABSTOP As Long = &H80
Const DT_TOP As Long = &H0
Const DT_VCENTER As Long = &H4
Const DT_WORDBREAK As Long = &H10
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'the actual text to scroll. This could also be loaded in from a text file
Dim ScrollText As String
Public EndingFlag As Boolean

Private Sub RunMain()
ScrollText = "H2O Networks  Snowman Media ilxz" & vbCrLf & vbCrLf & "Edition   4.06.9142002" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "策划   黄思夏" & vbCrLf & vbCrLf & "制作   黄思夏" & vbCrLf & vbCrLf & "编程   黄思夏" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "启动音乐   梁  睿" & vbCrLf & vbCrLf & vbCrLf & "控件提供   微软(Microsoft)中国            " & vbCrLf & vbCrLf & "Macromedia          " & vbCrLf & vbCrLf & "RealNetworks        " & vbCrLf & vbCrLf & "activePower         " & vbCrLf & vbCrLf & "Gamesman            " & vbCrLf & vbCrLf & "           武汉华工力学硕士研究生 - 刘玉锋" & vbCrLf & vbCrLf & "热情软件屋 - 李  海 " & vbCrLf & vbCrLf & vbCrLf & _
             "重点测试   清远市第一中学 - 信息科组" & vbCrLf & vbCrLf & _
             "      清远市第一中学 - PCC" & vbCrLf & vbCrLf & _
             "      清远市数码世界电脑城" & vbCrLf & vbCrLf & _
             "清远市勇创电脑" & vbCrLf & vbCrLf & vbCrLf & _
             "提供下载   太平洋电脑网" & vbCrLf & vbCrLf & "         华军软件园" & vbCrLf & vbCrLf & vbCrLf & _
             "特别鸣谢   朱素瑶" & vbCrLf & vbCrLf & "           何芝韵" & vbCrLf & vbCrLf & "           梁  睿" & vbCrLf & vbCrLf & "           岳  校" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                  "Thanks for using!" & vbCrLf & vbCrLf & vbCrLf & _
             "Copyright (C) 2000-2002 H2O Networks" & vbCrLf & vbCrLf & _
                                   "mailto:leask@21cn.com" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                             "Professional by Leask" & vbCrLf & vbCrLf & _
             "9.14.2002"
Dim LastFrameTime As Long
Const IntervalTime As Long = 20
Dim rt As Long
Dim DrawingRect As RECT
Dim UpperX As Long, UpperY As Long 'Upper left point of drawing rect
Dim RectHeight As Long
'show the form
Form2.Refresh
'Get the size of the drawing rectangle by suppying the DT_CALCRECT constant
rt = DrawText(Picture1.hdc, ScrollText, -1, DrawingRect, DT_CALCRECT)
If rt = 0 Then 'err
    EndingFlag = True
Else
    DrawingRect.Top = Picture1.ScaleHeight
    DrawingRect.Left = 0
    DrawingRect.Right = Picture1.ScaleWidth
    'Store the height of The rect
    RectHeight = DrawingRect.Bottom
    DrawingRect.Bottom = DrawingRect.Bottom + Picture1.ScaleHeight
End If
Do While Not EndingFlag
    If GetTickCount() - LastFrameTime > IntervalTime Then
        Picture1.Cls
        DrawText Picture1.hdc, ScrollText, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
        'update the coordinates of the rectangle
        DrawingRect.Top = DrawingRect.Top - 1
        DrawingRect.Bottom = DrawingRect.Bottom - 1
        'control the scolling and reset if it goes out of bounds
        If DrawingRect.Top < -(RectHeight) + 165 Then 'time to reset
               Exit Sub
            'DrawingRect.Top = Picture1.ScaleHeight
            'DrawingRect.Bottom = RectHeight + Picture1.ScaleHeight
        End If
        Picture1.Refresh
        LastFrameTime = GetTickCount()
    End If
   DoEvents
  Loop
Ext = True
End Sub

Private Sub Form_Load()
Picture1.Picture = LoadPicture(App.Path + "\SmM_Icos\logo.jpg")
Form1.Ly.MakeTop Form1, False
Form1.Enabled = False
Form1.Ly.MakeTop Me, True
Ext = False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   EndingFlag = True
    Ext = True
End Sub

Private Sub Form_Resize()
RunMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Form1.a00308.Checked = True Then Form1.Ly.MakeTop Form1, True
Form1.Enabled = True
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   EndingFlag = True
    Ext = True
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
If Ext = True Then
Unload Me
Set Form2 = Nothing
End If
End Sub
