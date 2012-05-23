VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Snowman Mdeia"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   6090
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&Y)"
      Height          =   285
      Left            =   4185
      TabIndex        =   2
      Top             =   2205
      Width           =   1860
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   585
      Picture         =   "Form4.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   45
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Height          =   3165
      Left            =   -2160
      Picture         =   "Form4.frx":0884
      ScaleHeight     =   3105
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   -1035
      Width           =   3795
   End
   Begin VB.Frame Frame1 
      Caption         =   "H2ontLeask  Snowman Media 1.0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   1710
      TabIndex        =   3
      Top             =   0
      Width           =   4335
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "版权所有 Copyleft 2000-2001  流动网络H2ont"
         Height          =   180
         Left            =   405
         TabIndex        =   8
         Top             =   1575
         Width           =   3780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "清远市第一中学第十四届艺术节特别版"
         Height          =   180
         Left            =   405
         TabIndex        =   7
         Top             =   585
         Width           =   3060
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "纯32位多功能播放媒体软件"
         Height          =   180
         Left            =   405
         TabIndex        =   6
         Top             =   360
         Width           =   2160
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "www.h2ont.com leask@china.com"
         Height          =   180
         Left            =   1575
         TabIndex        =   5
         Top             =   1800
         Width           =   2610
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "版本 HLS-SMM V1.0 E.QYA"
         Height          =   180
         Left            =   405
         TabIndex        =   4
         Top             =   810
         Width           =   2070
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本软件由  黄思夏 Leask  提供并保留最终解释权"
      Height          =   180
      Left            =   45
      TabIndex        =   9
      Top             =   2250
      Width           =   3960
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

