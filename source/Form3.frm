VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sorry - Snowman Mdeia"
   ClientHeight    =   2115
   ClientLeft      =   585
   ClientTop       =   90
   ClientWidth     =   6495
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&Y)"
      Height          =   330
      Left            =   4590
      TabIndex        =   6
      Top             =   1755
      Width           =   1860
   End
   Begin VB.Frame Frame1 
      Caption         =   "尊贵的用户，对不起："
      Height          =   1590
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   6405
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "电邮：leask@china.com"
         Height          =   180
         Left            =   4365
         TabIndex        =   5
         Top             =   1080
         Width           =   1890
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "传真：(501) 325-3048"
         Height          =   180
         Left            =   4365
         TabIndex        =   4
         Top             =   1305
         Width           =   1800
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "黄思夏 Leask"
         Height          =   180
         Left            =   4365
         TabIndex        =   3
         Top             =   855
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "如果在使用  Snowman Media  的过程中遇到任何问题欢迎与我联系。"
         Height          =   180
         Left            =   315
         TabIndex        =   2
         Top             =   450
         Width           =   5490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "  在  Snowman Media  这个版本中暂时不提供帮助文件。"
         Height          =   180
         Left            =   135
         TabIndex        =   1
         Top             =   225
         Width           =   4590
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Unload Me
End Sub
