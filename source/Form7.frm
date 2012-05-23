VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Begin VB.Form Form7 
   Caption         =   "Snowman Media  Flash Playing Window"
   ClientHeight    =   3390
   ClientLeft      =   3540
   ClientTop       =   3285
   ClientWidth     =   4515
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   3390
   ScaleWidth      =   4515
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   3390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4515
      _cx             =   4202268
      _cy             =   4200284
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
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ShockwaveFlash1.Height = Form7.Height
ShockwaveFlash1.Width = Form7.Width
End Sub

Private Sub form_resize()
ShockwaveFlash1.Height = Form7.Height
ShockwaveFlash1.Width = Form7.Width - 110
End Sub


