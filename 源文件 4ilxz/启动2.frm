VERSION 5.00
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "SmM_Tools.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   Enabled         =   0   'False
   Icon            =   "启动2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3960
      Top             =   2070
   End
   Begin API控制大全.LyfTools Ly 
      Left            =   4725
      Top             =   2025
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim cli As Integer
Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance = True Then End
Me.Picture = LoadPicture(App.Path + "\SmM_Icos\logo.jpg")
If Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ShowPic") = True And Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Sting") = False Then
Me.Show
Me.Ly.MakeTop Me, True
End If
i = 0

End Sub


Private Sub Timer1_Timer()
On Error Resume Next

If i = 0 Then
 If Len(Command) > 0 Then
     If Right(Command, 1) = """" Then
      For cli = 1 To Len(Command)
       If Mid(Command, cli, 1) = """" Then
          myWriteINI App.Path + "\SmM_Start.dat", "Start", "PlayFile", Mid(Command, cli + 1, Len(Command) - cli - 1)
    
        Exit For
        End If
    Next
    Else
                 myWriteINI App.Path + "\SmM_Start.dat", "Start", "PlayFile", Command
  
    End If
           Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "PlayFile", True
Else:
   Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "PlayFile", False
End If

Dim DriveType As Integer
Dim rtn As String
Dim AllDrives As String
Dim JustOneDrive As String
Dim CDRom As String
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
If Len(CDRom) > 0 Then Ly.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "CDrom", CDRom
Shell (App.Path + "\SmM_Player.exe"), vbNormalFocus
End If



i = i + 1
If i > 2 Then End
End Sub

