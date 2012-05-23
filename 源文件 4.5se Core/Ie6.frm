VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "SmM_Tools.ocx"
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "SmM_Snowflake.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "S.m.M. Browser"
   ClientHeight    =   5985
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6045
   Icon            =   "Ie6.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   6045
   Begin ACTIVESKINLibCtl.SkinForm Skin 
      Height          =   480
      Left            =   5265
      OleObjectBlob   =   "Ie6.frx":2CFA
      TabIndex        =   1
      Top             =   6660
      Width           =   480
   End
   Begin API¿ØÖÆ´óÈ«.LyfTools Ly 
      Left            =   5895
      Top             =   6705
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6615
      Top             =   6705
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   6000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      ExtentX         =   10583
      ExtentY         =   10583
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
If App.PrevInstance = True Then End
Skin.SkinPath = App.Path + "\SmM_Skin"
Me.Top = 0
Me.Left = Screen.Width - Me.Width
Me.Height = Screen.Height - 400
Me.Show
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Width < 5560 Then Me.Width = 5560
If Me.Height < 6000 Then Me.Height = 6000

Web.Width = Me.Width - 325
Web.Height = Me.Height - 460
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
'On Error Resume Next
If Ly.GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "NetShow") = True Then
Web.Navigate Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "NetFile")
Ly.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "NetShow", False
End If

End Sub
