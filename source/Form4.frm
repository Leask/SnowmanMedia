VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About  Snowman Mdeia  2.0"
   ClientHeight    =   3195
   ClientLeft      =   2880
   ClientTop       =   3120
   ClientWidth     =   6810
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "H2ontLeask  Snowman Media 2.0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4335
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "版权所有 Copyleft 2000-2001  流动网络H2ont"
         Height          =   180
         Left            =   450
         TabIndex        =   7
         Top             =   900
         Width           =   3780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "中华人民共和国青少年计算机制作比赛特别版"
         Height          =   180
         Left            =   450
         TabIndex        =   6
         Top             =   450
         Width           =   3600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "纯32位多功能播放媒体软件"
         Height          =   180
         Left            =   450
         TabIndex        =   5
         Top             =   270
         Width           =   2160
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "www.h2ont.com h2ont@china.com"
         Height          =   180
         Left            =   1620
         TabIndex        =   4
         Top             =   1080
         Width           =   2610
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "版本 HLS-SMM V2.0.2 C.TBB"
         Height          =   180
         Left            =   450
         TabIndex        =   3
         Top             =   630
         Width           =   2250
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   1815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "Form4.frx":1582
      Top             =   1350
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&Y)"
      Height          =   285
      Left            =   4365
      TabIndex        =   1
      Top             =   2880
      Width           =   2445
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3735
      Left            =   4365
      Picture         =   "Form4.frx":1A8E
      ScaleHeight     =   3675
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   -900
      Width           =   6060
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Unload Me
End Sub

