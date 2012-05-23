VERSION 5.00
Begin VB.Form CDRom 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CD-Rom Detect"
   ClientHeight    =   1395
   ClientLeft      =   3120
   ClientTop       =   3225
   ClientWidth     =   3060
   Icon            =   "CD-Rom Detect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1395
   ScaleWidth      =   3060
   Begin VB.CommandButton Check 
      Caption         =   "&Check for CD-Rom Dirve Letter..."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox CDLetter 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Your Current CD-Rom Drive is:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "CDRom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


