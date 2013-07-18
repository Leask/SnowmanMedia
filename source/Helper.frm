VERSION 5.00
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "SmM_Tools.ocx"
Object = "{8106FB53-4B58-11D3-AFB6-DC8009C10000}#3.0#0"; "SmM_CDCheck.ocx"
Object = "{CFB094A6-8FF0-4EF7-A644-ED122CC38E57}#1.0#0"; "SmM_Tray.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "S.m.M. Helper"
   ClientHeight    =   720
   ClientLeft      =   3870
   ClientTop       =   3300
   ClientWidth     =   3315
   Enabled         =   0   'False
   Icon            =   "Helper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin TASKICONLib.TaskIcon Ti 
      Left            =   1710
      Top             =   45
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "Helper.frx":2CFA
      ToolTipText     =   "欢迎使用 Snowman Media ilxz 4!"
      TitleText       =   ""
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   2250
      Pattern         =   "*.dat;*.cda"
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   600
   End
   Begin CDNotification.CDNotify CD 
      Left            =   1125
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   585
      Top             =   90
   End
   Begin API控制大全.LyfTools Ly 
      Left            =   90
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Menu frthrt 
      Caption         =   "a"
      Begin VB.Menu dfhgg 
         Caption         =   "启动 Snowman Media ilxz(&S)"
      End
      Begin VB.Menu gregdfgfdg 
         Caption         =   "-"
      End
      Begin VB.Menu vrg 
         Caption         =   "选项(&O)"
      End
      Begin VB.Menu grvdcvasdf 
         Caption         =   "关联媒体(&B)"
      End
      Begin VB.Menu asftg 
         Caption         =   "-"
      End
      Begin VB.Menu vewrg 
         Caption         =   "帮助(&H)"
      End
      Begin VB.Menu vrfr 
         Caption         =   "-"
      End
      Begin VB.Menu dfgdsfg 
         Caption         =   "访问流动网络(&V)"
      End
      Begin VB.Menu dfsgs 
         Caption         =   "交流和反馈(&C)"
      End
      Begin VB.Menu sdfvgg 
         Caption         =   "-"
      End
      Begin VB.Menu zxcweftg 
         Caption         =   "禁用助手(&D)"
      End
      Begin VB.Menu vr 
         Caption         =   "关闭助手(&T)"
      End
      Begin VB.Menu sdf 
         Caption         =   "-"
      End
      Begin VB.Menu sdfsdf 
         Caption         =   "退出菜单(&X)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CDRom As String

Private Sub CD_Arrival(ByVal Drive As String)
Dim i As Long
If Ly.FileExists(CDRom + "MPEGAV\AVSEQ01.DAT") = True Or Ly.FileExists(CDRom + "MPEGAV\MUSIC01.DAT") = True Then
File1.Path = CDRom + "MPEGAV"
Open App.Path + "\SmM_List.sml" For Output As #1
    For i = 0 To File1.ListCount - 1
     Print #1, CDRom + "MPEGAV\" + File1.List(i)
    Next i
   Close (1)
Shell (App.Path + "\Snowman.exe " + App.Path + "\SmM_List.sml")
Exit Sub
End If
If Ly.FileExists(CDRom + "Track01.cda") = True Then
File1.Path = CDRom
Open App.Path + "\SmM_List.sml" For Output As #1
    For i = 0 To File1.ListCount - 1
     Print #1, CDRom + File1.List(i)
    Next i
   Close (1)
Shell (App.Path + "\Snowman.exe " + App.Path + "\SmM_List.sml")
End If

End Sub

Private Sub dfgdsfg_Click()
Ly.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "NetShow", True
Ly.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "NetFile", "http://www.51.net"
Shell App.Path + "\SmM_IntBrowser.exe", vbMinimizedFocus

End Sub

Private Sub dfhgg_Click()
Shell (App.Path + "\Snowman.exe"), vbNormalFocus
End Sub

Private Sub dfsgs_Click()
Ly.SendMail "leask@21cn.com"
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then End
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "StartUp") = True Then Ti.ShowIcon
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoMedia") = True Then Shell (App.Path + "\SmM_Types.exe")
Dim DriveType As Long
Dim rtn As String
Dim AllDrives As String
Dim JustOneDrive As String
AllDrives = Space$(64)

rtn = GetLogicalDriveStrings(Len(AllDrives), AllDrives) 'call the function to get the string containing all drives
AllDrives = Left(AllDrives, rtn) 'trim off trailing chr(0)'s.  AllDrives$ now contains all the drive letters.

Do
  rtn = InStr(AllDrives, Chr(0)) 'find the first separating chr(0)
  If rtn Then 'if there is one then
     JustOneDrive = Left(AllDrives, rtn) 'extract the drive up to the chr(0)
     AllDrives = Mid(AllDrives, rtn + 1, Len(AllDrives)) 'and remove that from the Alldrives string, so it won't be checked again
     
     rtn = GetDriveType(JustOneDrive) 'check what drive it is
     If rtn = DRIVE_CDROM Then 'if it is a CD-Rom drive then
        CDRom = Left(UCase(JustOneDrive), 3) 'return the drive letter to the user
        Exit Do
     End If
  End If
Loop Until AllDrives = "" Or DriveType = DRIVE_CDROM

End Sub

Private Sub Form_Unload(Cancel As Integer)

End
End Sub


Private Sub grvdcvasdf_Click()

Shell (App.Path + "\SmM_Types.exe")
End Sub


Private Sub Ti_OnIconEvent(ByVal EventMsg As Long)
If EventMsg = 515 Then
Shell (App.Path + "\Snowman.exe"), vbNormalFocus
Unload Me
End If
If EventMsg = 516 Then
PopupMenu Me.frthrt, 40, Screen.Width, Screen.Height
End If
If EventMsg = 519 Then
Ti.ShowMsg "欢迎使用 Snowman Media ilxz 4!", ICON_INFO, "Snowman Media", 0
End If
End Sub

Private Sub Timer1_Timer()
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "StartUp") = False Then Unload Me
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Sting") = True Then Unload Me
End Sub

Private Sub vewrg_Click()
Shell (App.Path + "\SmM_Help.exe"), vbNormalFocus

End Sub

Private Sub vr_Click()

Unload Me
End Sub

Private Sub vrg_Click()

Shell (App.Path + "\SmM_Settings.exe"), vbNormalFocus
End Sub

Private Sub zxcweftg_Click()

Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "StartUp", False
Unload Me
End Sub
