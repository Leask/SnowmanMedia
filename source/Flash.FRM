VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "LYFTOOLS.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form7 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Sm.M. Flash Player"
   ClientHeight    =   4515
   ClientLeft      =   3540
   ClientTop       =   3285
   ClientWidth     =   6765
   Icon            =   "Flash.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   4515
   ScaleWidth      =   6765
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3600
      Top             =   1575
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin API控制大全.LyfTools LyfTools1 
      Left            =   2565
      Top             =   945
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin ACTIVESKINLibCtl.SkinForm SkinForm1 
      Height          =   480
      Left            =   4185
      OleObjectBlob   =   "Flash.frx":030A
      TabIndex        =   1
      Top             =   3195
      Width           =   480
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   4515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6765
      _cx             =   4206237
      _cy             =   4202268
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.Label Label1 
      Caption         =   """"
      Height          =   510
      Left            =   2340
      TabIndex        =   2
      Top             =   4005
      Width           =   1635
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Filename As String
Private Sub Form_Load()
 SkinForm1.SkinPath = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path")
If Len(Command) > 0 Then
 Filename = sReplace(Command, Label1.Caption, "")
 Filename = sReplace(Filename, Label1.Caption, "")
ShockwaveFlash1.Movie = Filename
Else
   CommonDialog1.Filename = ""
        CommonDialog1.Filter = "Flash媒体:Swf" & _
      "|*.swf|所有文件:*.*|*.*"
        CommonDialog1.FilterIndex = 1
    CommonDialog1.ShowOpen
    If Len(CommonDialog1.Filename) > 0 Then
  ShockwaveFlash1.Movie = CommonDialog1.Filename
     Else: Unload Me
      End If
End If
Me.Caption = Me.ShockwaveFlash1.Movie + " - Sm.M. Flash Player"

End Sub
Function sReplace(SearchLine As String, SearchFor As String, ReplaceWith As String)
Dim vSearchLine As String, found As Integer

found = InStr(SearchLine, SearchFor): vSearchLine = SearchLine
If found <> 0 Then
vSearchLine = ""
If found > 1 Then vSearchLine = Left(SearchLine, found - 1)
vSearchLine = vSearchLine + ReplaceWith
If found + Len(SearchFor) - 1 < Len(SearchLine) Then _
vSearchLine = vSearchLine + Right$(SearchLine, Len(SearchLine) - found - Len(SearchFor) + 1)
End If
sReplace = vSearchLine

End Function

Private Sub Form_Resize()
 On Error Resume Next
 ShockwaveFlash1.Width = Me.Width - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "w", "") - 6240) - 120
 ShockwaveFlash1.Height = Me.Height - (myReadINI(SkinForm1.SkinPath + "\skin_info.skin", "FORM", "h", "") - 6345) - 405
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
