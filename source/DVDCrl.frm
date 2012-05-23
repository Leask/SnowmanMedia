VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   0  'None
   Caption         =   "Snowman Media  3.0"
   ClientHeight    =   570
   ClientLeft      =   4260
   ClientTop       =   -540
   ClientWidth     =   3480
   ControlBox      =   0   'False
   Icon            =   "DVDCrl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   570
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   510
      Left            =   0
      TabIndex        =   2
      Top             =   495
      Width           =   12075
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3105
      ScaleHeight     =   165
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   90
      Width           =   195
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "§ç"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   240
      End
   End
   Begin SmMDVDPlayer.Form_TaskBar Form_TaskBar1 
      Left            =   -45
      Top             =   360
      _ExtentX        =   1693
      _ExtentY        =   423
   End
   Begin VB.Image Image9 
      Height          =   285
      Left            =   2250
      Picture         =   "DVDCrl.frx":030A
      Stretch         =   -1  'True
      Top             =   45
      Width           =   300
   End
   Begin VB.Image Image8 
      Height          =   285
      Left            =   2610
      Picture         =   "DVDCrl.frx":0709
      Stretch         =   -1  'True
      Top             =   135
      Width           =   300
   End
   Begin VB.Image Image7 
      Height          =   285
      Left            =   1890
      Picture         =   "DVDCrl.frx":0B08
      Stretch         =   -1  'True
      Top             =   135
      Width           =   300
   End
   Begin VB.Image Image6 
      Height          =   285
      Left            =   1530
      Picture         =   "DVDCrl.frx":0F07
      Stretch         =   -1  'True
      Top             =   45
      Width           =   300
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   810
      Picture         =   "DVDCrl.frx":1306
      Stretch         =   -1  'True
      Top             =   45
      Width           =   300
   End
   Begin VB.Image Image4 
      Height          =   285
      Left            =   1170
      Picture         =   "DVDCrl.frx":1705
      Stretch         =   -1  'True
      Top             =   135
      Width           =   300
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   495
      Picture         =   "DVDCrl.frx":1B04
      Stretch         =   -1  'True
      Top             =   180
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   135
      Picture         =   "DVDCrl.frx":1F03
      Stretch         =   -1  'True
      Top             =   45
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   135
      Picture         =   "DVDCrl.frx":2302
      Top             =   45
      Width           =   4515
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Image2_Click()
On Error Resume Next
Form4.DVD.Play
End Sub

Private Sub Image3_Click()
On Error Resume Next
Form4.DVD.Pause
End Sub

Private Sub Image4_Click()
On Error Resume Next
Form4.DVD.PlayPrevChapter
End Sub

Private Sub Image5_Click()
On Error Resume Next
Form4.DVD.PlayBackwards (5)
End Sub

Private Sub Image6_Click()
On Error Resume Next
Form4.DVD.PlayForwards (5)
End Sub

Private Sub Image7_Click()
On Error Resume Next
Form4.DVD.PlayNextChapter
End Sub
Private Sub Image8_Click()
On Error Resume Next
Form4.DVD.Eject
End Sub
Private Sub Image9_Click()
On Error Resume Next
Form4.DVD.Stop
End Sub

Private Sub Label1_Click()
Unload Form4
Unload Me
End Sub

