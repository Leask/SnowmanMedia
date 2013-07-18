VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于 S.m.M Video Capturer"
   ClientHeight    =   1500
   ClientLeft      =   1665
   ClientTop       =   3420
   ClientWidth     =   6105
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&Y)"
      Height          =   330
      Left            =   4590
      TabIndex        =   0
      Top             =   315
      Width           =   1230
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "版本: ilxz 2.76.2002"
      Height          =   225
      Index           =   2
      Left            =   1980
      TabIndex        =   3
      Top             =   585
      Width           =   1875
   End
   Begin VB.Shape Shape1 
      Height          =   780
      Left            =   225
      Top             =   315
      Width           =   780
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   750
      Left            =   225
      Picture         =   "About.frx":000C
      Top             =   315
      Width           =   750
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "Snowman Media Video Capturer"
      Height          =   225
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   270
      Width           =   2820
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "Copyright (C) 2000-2002 H2O Networks"
      Height          =   225
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   1035
      Width           =   3405
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()

End Sub
