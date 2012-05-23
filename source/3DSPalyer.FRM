VERSION 5.00
Object = "{481BB642-0766-11D2-9830-444553540000}#1.0#0"; "MASFLC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "LYFTOOLS.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Snowman Media 3DS Player  1.0"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6765
   Icon            =   "3DSPalyer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   6765
   StartUpPosition =   2  '屏幕中心
   Begin API控制大全.LyfTools LyfTools1 
      Left            =   3510
      Top             =   2430
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2250
      Top             =   1035
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinForm SkinForm1 
      Height          =   480
      Left            =   2205
      OleObjectBlob   =   "3DSPalyer.frx":1582
      TabIndex        =   1
      Top             =   1890
      Width           =   480
   End
   Begin MASFLCLibCtl.MasFLCPlayer MFCI1 
      Height          =   4515
      Left            =   0
      OleObjectBlob   =   "3DSPalyer.frx":15CB
      TabIndex        =   0
      Top             =   0
      Width           =   6765
   End
   Begin VB.Label Label1 
      Caption         =   """"
      Height          =   465
      Left            =   4275
      TabIndex        =   2
      Top             =   3870
      Width           =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Filename As String
Private Sub Form_Load()
 SkinForm1.SkinPath = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media  3.0", "Skin_Path")
If Len(Command) > 0 Then
 Filename = sReplace(Command, Label1.Caption, "")
 Filename = sReplace(Filename, Label1.Caption, "")
 MFCI1.Filename = Filename
 Else
   CommonDialog1.Filename = ""
        CommonDialog1.Filter = "3DS媒体:Flx" & _
      "|*.fl*|所有文件:*.*|*.*"
        CommonDialog1.FilterIndex = 1
    CommonDialog1.ShowOpen
    If Len(CommonDialog1.Filename) > 0 Then
  MFCI1.Filename = CommonDialog1.Filename
     Else: Unload Me
      End If
End If
End Sub

Private Sub Form_Resize()
MFCI1.Height = Me.Height - 510
MFCI1.Width = Me.Width - 210
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


