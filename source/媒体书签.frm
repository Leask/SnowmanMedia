VERSION 5.00
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sm.M. B.M."
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   Icon            =   "媒体书签.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4875
   StartUpPosition =   2  '屏幕中心
   Begin ACTIVESKINLibCtl.SkinForm SkinForm1 
      Height          =   480
      Left            =   4500
      OleObjectBlob   =   "媒体书签.frx":1582
      TabIndex        =   0
      Top             =   3825
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   2310
      Left            =   -405
      TabIndex        =   1
      Top             =   -45
      Width           =   6360
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "媒体书签 [ B ]"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   945
         TabIndex        =   5
         Top             =   675
         Width           =   1680
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "媒体书签 [ D ]"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2295
         TabIndex        =   4
         Top             =   1350
         Width           =   1905
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "媒体书签 [ A ]"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   945
         TabIndex        =   3
         Top             =   360
         Width           =   1680
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "媒体书签 [ C ]"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2295
         TabIndex        =   2
         Top             =   1035
         Width           =   1905
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H000000C0&
         BorderStyle     =   2  'Dash
         Height          =   915
         Left            =   540
         Shape           =   1  'Square
         Top             =   1125
         Width           =   735
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   3555
         Picture         =   "媒体书签.frx":15CB
         Top             =   315
         Width           =   480
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00800000&
         Height          =   735
         Left            =   3510
         Top             =   945
         Width           =   1635
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H0080FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         BorderStyle     =   3  'Dot
         Height          =   1635
         Left            =   765
         Shape           =   4  'Rounded Rectangle
         Top             =   180
         Width           =   3570
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   1725
         Left            =   720
         Shape           =   4  'Rounded Rectangle
         Top             =   135
         Width           =   3660
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
On Error Resume Next
 SkinForm1.SkinPath = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path")

End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set Form5 = Nothing
End Sub

Private Sub Option1_Click()
On Error Resume Next
If Len(Form102.MediaPlayer1.Filename) > 0 Then
Form102.LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_A", Form102.MediaPlayer1.Filename
Form102.LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_A", Form102.MediaPlayer1.CurrentPosition
Unload Me
Exit Sub
End If
If Len(Form102.RA.Source) > 0 Then
Form102.LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_A", Form102.sReplace(Form102.RA.Source, "file://", "")
Form102.LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_A", Form102.RA.GetPosition
End If
Unload Me
End Sub
Private Sub Option2_Click()
On Error Resume Next
If Len(Form102.MediaPlayer1.Filename) > 0 Then
Form102.LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_B", Form102.MediaPlayer1.Filename
Form102.LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_B", Form102.MediaPlayer1.CurrentPosition
Unload Me
Exit Sub
End If
If Len(Form102.RA.Source) > 0 Then
Form102.LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_B", Form102.sReplace(Form102.RA.Source, "file://", "")
Form102.LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_B", Form102.RA.GetPosition
End If
Unload Me
End Sub
Private Sub Option3_Click()
On Error Resume Next
If Len(Form102.MediaPlayer1.Filename) > 0 Then
Form102.LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_C", Form102.MediaPlayer1.Filename
Form102.LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_C", Form102.MediaPlayer1.CurrentPosition
Unload Me
Exit Sub
End If
If Len(Form102.RA.Source) > 0 Then
Form102.LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_C", Form102.sReplace(Form102.RA.Source, "file://", "")
Form102.LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_C", Form102.RA.GetPosition
End If
Unload Me
End Sub
Private Sub Option4_Click()
On Error Resume Next
If Len(Form102.MediaPlayer1.Filename) > 0 Then
Form102.LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_D", Form102.MediaPlayer1.Filename
Form102.LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_D", Form102.MediaPlayer1.CurrentPosition
Unload Me
Exit Sub
End If
If Len(Form102.RA.Source) > 0 Then
Form102.LyfTools1.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Name_D", Form102.sReplace(Form102.RA.Source, "file://", "")
Form102.LyfTools1.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Alo_Rute_D", Form102.RA.GetPosition
End If
Unload Me
End Sub
