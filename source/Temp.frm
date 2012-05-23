VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   5085
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5880
   LinkTopic       =   "Form6"
   ScaleHeight     =   5085
   ScaleWidth      =   5880
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   3645
      Width           =   1995
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   4500
      TabIndex        =   0
      Top             =   4815
      Width           =   2085
      Begin VB.PictureBox MediaPlayer1 
         Height          =   1095
         Left            =   1620
         ScaleHeight     =   1035
         ScaleWidth      =   2025
         TabIndex        =   1
         Top             =   540
         Width           =   2085
      End
   End
   Begin VB.Label Label1 
      Caption         =   "ControlPanel,StatusBar"
      Height          =   375
      Left            =   3105
      TabIndex        =   3
      Top             =   2250
      Width           =   2625
   End
   Begin VB.Image Image2 
      Height          =   2700
      Left            =   0
      Picture         =   "TEMP.frx":0000
      Top             =   0
      Width           =   2700
   End
   Begin VB.Image Image1 
      Height          =   2130
      Left            =   5400
      Picture         =   "TEMP.frx":17BF2
      Top             =   4545
      Width           =   3525
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
