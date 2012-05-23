VERSION 5.00
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sm.M. A.S.D"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "自动关机.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5445
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4590
      Top             =   6615
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   5460
      Left            =   -270
      TabIndex        =   0
      Top             =   -135
      Width           =   6765
      Begin VB.Timer Timer3 
         Interval        =   1000
         Left            =   3735
         Top             =   1530
      End
      Begin VB.Timer Timer2 
         Interval        =   2000
         Left            =   4275
         Top             =   1530
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   4770
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "自动关机.frx":1582
         Top             =   495
         Width           =   420
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "关闭自动关机功能"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   675
         TabIndex        =   5
         Top             =   1710
         Width           =   1770
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "指定时间倒数完毕后自动关机"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   675
         TabIndex        =   4
         Top             =   1350
         Width           =   2715
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "同步列表播放完毕后自动关机"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   675
         TabIndex        =   3
         Top             =   990
         Value           =   -1  'True
         Width           =   2715
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3555
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "自动关机.frx":1587
         Top             =   495
         Width           =   420
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   4140
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "自动关机.frx":158C
         Top             =   495
         Width           =   420
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   900
         Left            =   4860
         Picture         =   "自动关机.frx":1591
         Top             =   1035
         Width           =   660
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         Height          =   330
         Left            =   630
         Shape           =   4  'Rounded Rectangle
         Top             =   1665
         Width           =   2490
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         Height          =   330
         Left            =   630
         Shape           =   4  'Rounded Rectangle
         Top             =   1305
         Width           =   2850
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         Height          =   330
         Left            =   630
         Shape           =   4  'Rounded Rectangle
         Top             =   945
         Width           =   3120
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   3930
         Left            =   405
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   ":      :"
         ForeColor       =   &H008080FF&
         Height          =   600
         Left            =   4005
         TabIndex        =   8
         Top             =   495
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   675
         MouseIcon       =   "自动关机.frx":198A
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   270
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   330
         Left            =   450
         Shape           =   4  'Rounded Rectangle
         Top             =   225
         Width           =   1590
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000080FF&
         BorderStyle     =   2  'Dash
         Height          =   420
         Left            =   3375
         Top             =   405
         Width           =   1995
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         Height          =   375
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   225
         Width           =   5190
      End
   End
   Begin ACTIVESKINLibCtl.SkinForm SkinForm1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "自动关机.frx":1ADC
      TabIndex        =   9
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const EWX_SHUTDOWN As Long = 1
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Dim Hrs
Dim Mnt
Dim AMPM
Dim SetAlarm
Dim Hours As Integer
Dim Minutes As Integer
Dim Seconds As Integer
Dim Time As Date
Public Asd As Boolean
    Dim lngResult As Long
    Dim Ta As String, Tb As String
Private Sub Mydisplay()
    Hours = Val(Text1.Text)
    Minutes = Val(Text2.Text)
    Seconds = Val(Text3.Text)
    Time = TimeSerial(Hours, Minutes, Seconds)
    Label1.Caption = Format$(Time, "hh") & ":" & Format$(Time, "nn") & ":" & Format$(Time, "ss")
End Sub






Private Sub Form_Load()
      Asd = False
 SkinForm1.SkinPath = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path")
    Timer2.Enabled = False
    Timer1.Enabled = False
    Hours = 0
    Minutes = 0
    Seconds = 0
    Time = 0
    Ta = "00:00:00"
    Tb = Ta
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set Form4 = Nothing
End Sub

Private Sub Label1_Click()
If Timer1.Enabled = False Then
Timer1.Enabled = True
Timer3.Enabled = True
MsgBox ("已经恢复计时.")
Exit Sub
End If
If Timer1.Enabled = True Then
Timer1.Enabled = False
Timer3.Enabled = False
MsgBox ("已经停止计时.")
Exit Sub
End If
End Sub

Private Sub Option1_Click()
Label1.Enabled = False
End Sub

Private Sub Option2_Click()
Label1.Enabled = True
End Sub

Private Sub Option3_Click()
Unload Me
End Sub
Private Sub Text1_Change()
    Mydisplay

End Sub
Private Sub Text2_Change()
    Mydisplay

End Sub
Private Sub Text3_Change()
    Mydisplay

End Sub
Private Sub Timer1_Timer()
    If Timer3.Enabled = False Then
    Timer1.Enabled = False
    Exit Sub
    End If
     Tb = Ta
     Ta = Label1.Caption
    Timer1.Enabled = False
    If (Format$(Time, "hh") & ":" & Format$(Time, "nn") & ":" & Format$(Time, "ss")) <> "00:00:00" Then 'Counter to continue loop until 0
        Time = DateAdd("s", -1, Time)
        Label1.Visible = False
        Label1.Caption = Format$(Time, "hh") & ":" & Format$(Time, "nn") & ":" & Format$(Time, "ss")
        Label1.Visible = True
        Timer1.Enabled = True
    Else
        Timer1.Enabled = False
        If Ta <> Tb And Option2.Value = True Then
     lngResult = ExitWindowsEx(EWX_SHUTDOWN, 0&)
         End If
    End If
End Sub

Private Sub Timer2_Timer()
If Asd = True Then lngResult = ExitWindowsEx(EWX_SHUTDOWN, 0&)

End Sub

Private Sub Timer3_Timer()
If Option1.Value = True Then
Timer2.Enabled = True
Timer1.Enabled = False
End If
If Option2.Value = True Then
Timer2.Enabled = False
Timer1.Enabled = True
End If
End Sub

