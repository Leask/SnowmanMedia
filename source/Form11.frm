VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sorry - Snowman Mdeia"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   6990
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "尊贵的用户，对不起："
      Height          =   1590
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6990
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "  在  Snowman Media  这个版本中暂时不提供此项功能。"
         Height          =   180
         Left            =   135
         TabIndex        =   6
         Top             =   225
         Width           =   4590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "如果在使用  Snowman Media  的过程中遇到任何不便敬请原谅，欢迎与我联系。"
         Height          =   180
         Left            =   315
         TabIndex        =   5
         Top             =   450
         Width           =   6390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "黄思夏 Leask"
         Height          =   180
         Left            =   4815
         TabIndex        =   4
         Top             =   855
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "传真：(501) 325-3048"
         Height          =   180
         Left            =   4815
         TabIndex        =   3
         Top             =   1305
         Width           =   1800
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "电邮：leask@china.com"
         Height          =   180
         Left            =   4815
         TabIndex        =   2
         Top             =   1080
         Width           =   1890
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&Y)"
      Height          =   330
      Left            =   5085
      TabIndex        =   0
      Top             =   1710
      Width           =   1860
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

