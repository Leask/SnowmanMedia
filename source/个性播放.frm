VERSION 5.00
Begin VB.Form Form100 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Snowman Media  3.0"
   ClientHeight    =   4440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8220
   ForeColor       =   &H00000000&
   Icon            =   "个性播放.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   720
      TabIndex        =   0
      Top             =   1305
      Width           =   6855
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "www.h2ont.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   3825
         MouseIcon       =   "个性播放.frx":1582
         MousePointer    =   99  'Custom
         TabIndex        =   5
         ToolTipText     =   "点击访问 流动网络 H2ont"
         Top             =   450
         Width           =   1410
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "leask@21cn.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   5265
         MouseIcon       =   "个性播放.frx":16D4
         MousePointer    =   99  'Custom
         TabIndex        =   4
         ToolTipText     =   "点击联系作者"
         Top             =   450
         Width           =   1560
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright (C) 2000-2001 H2ont Leask  www.h2ont.com leask@21cn.com"
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
         Left            =   495
         TabIndex        =   3
         Top             =   450
         Width           =   6360
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Snowman Media ilxz 3.5"
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
         Height          =   465
         Left            =   2925
         TabIndex        =   2
         Top             =   45
         Width           =   3885
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "10.30.2001"
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
         Left            =   5895
         TabIndex        =   1
         Top             =   720
         Width           =   915
      End
   End
End
Attribute VB_Name = "Form100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim j As Integer
Dim ls As Integer
Dim i As Integer
Dim oldx As Integer
Dim oldy As Integer
Dim Coloury As Integer
Dim Snow(1000, 2) As Integer, Amounty As Integer
Public Et As Boolean



Private Sub Form_Click()
If Form102.f200 = 0 Then
Form102.WindowState = 0

Form102.Show
End If
End Sub
Private Sub Form_GotFocus()
If Form102.f200 = 0 Then
Form102.WindowState = 0
Form102.Show
End If
End Sub


Private Sub Form_Load()
Et = True
Form100.Show
Form100.Frame1.Left = Form100.Width - 6900
Form100.Frame1.Top = Form100.Height - 1050
Call SnowB
Form102.Show
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then Et = False
End Sub

Private Sub Form_Resize()
If Form102.f200 = 0 Then
Form102.WindowState = 0

Form102.Show
End If

End Sub


Sub SnowB()
DoEvents
Randomize: Amounty = 325
For j = 1 To Amounty
Snow(j, 0) = Int(Rnd * Form100.Width)
Snow(j, 1) = Int(Rnd * Form100.Height)
Snow(j, 2) = 10 + (Rnd * 20)
If Et = False Then Exit For
Next j
Do While Not (DoEvents = 0)
For ls = 1 To 10
For i = 1 To Amounty
oldx = Snow(i, 0): oldy = Snow(i, 1): Snow(i, 1) = Snow(i, 1) + Snow(i, 2)
If Snow(i, 1) > Form100.Height Then Snow(i, 1) = 0: Snow(i, 2) = 5 + (Rnd * 30): Snow(i, 0) = Int(Rnd * Form100.Width): oldx = 0: oldy = 0
Coloury = 8 * (Snow(i, 2) - 10): Coloury = 60 + Coloury: PSet (oldx, oldy), QBColor(0): PSet (Snow(i, 0), Snow(i, 1)), RGB(Coloury, Coloury, Coloury)
If Et = False Then Exit For
Next i
If Et = False Then Exit For
Next ls
Label1.Refresh
Label2.Refresh
Label3.Refresh
If Et = False Then Exit Do
Loop
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Form100 = Nothing
End Sub

Private Sub Label18_Click()
               Form102.LyfTools1.SendMail ("leask@21cn.com")
End Sub

Private Sub Label4_Click()
              If Form102.LyfTools1.IsConnected = True Then
              
              Form102.LyfTools1.HttpTo ("http://www.h2ont.com")
              Else
               MsgBox ("无法访问 流动网络 可能你的计算机尚未连接网络,请确认连接网络后重试.")
               End If

End Sub
