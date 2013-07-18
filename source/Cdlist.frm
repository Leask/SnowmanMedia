VERSION 5.00
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "SmM_Tools.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2160
   Enabled         =   0   'False
   Icon            =   "Cdlist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   2160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1110
      Left            =   45
      Pattern         =   "*.cda"
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   1500
   End
   Begin API控制大全.LyfTools Ly 
      Left            =   1710
      Top             =   315
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
Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance = True Then End
Dim i As Long
Dim DriveType As Long
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
If Len(CDRom) > 0 Then
If Ly.FileExists(CDRom + "Track01.cda") = True Then
File1.Path = CDRom
Open App.Path + "\SmM_List.sml" For Output As #1
    For i = 0 To File1.ListCount - 1
     Print #1, CDRom + File1.List(i)
    Next i
   Close (1)
Shell (App.Path + "\Snowman.exe " + App.Path + "\SmM_List.sml"), vbNormalFocus
End If
End If
End
End Sub
