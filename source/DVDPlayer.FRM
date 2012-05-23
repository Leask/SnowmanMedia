VERSION 5.00
Object = "{38EE5CE1-4B62-11D3-854F-00A0C9C898E7}#1.0#0"; "MSWEBDVD.DLL"
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Snowman Media DVD Player  1.0"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   Icon            =   "DVDPlayer.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   3390
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   WindowState     =   2  'Maximized
   Begin MSWEBDVDLibCtl.MSWebDVD DVD 
      Height          =   3390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4515
      _cx             =   7964
      _cy             =   5980
      DisableAutoMouseProcessing=   0   'False
      BackColor       =   1048592
      EnableResetOnStop=   0   'False
      ColorKey        =   0
      WindowlessActivation=   0   'False
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmTest.Show
End Sub

Private Sub Form_Resize()
DVD.Height = Me.Height
DVD.Width = Me.ScaleWidth
End Sub
