VERSION 5.00
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#4.0#0"; "LYFTOOLS.OCX"
Begin VB.Form Form1 
   Caption         =   "Mail To Us"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin 刘玉锋的VB超级工具集.LyfTools LyfTools1 
      Height          =   480
      Left            =   4050
      TabIndex        =   0
      Top             =   2610
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
LyfTools1.SendMail "h2ont@china.com"
End
End Sub
