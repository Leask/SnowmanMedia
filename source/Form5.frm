VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Medias Opening Window  - Snowman Media  2.0"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5460
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Caption         =   "����ý���ַ:"
      Height          =   600
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5415
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Top             =   225
         Width           =   5235
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�������е�ַ:"
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   630
      Width           =   2535
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "http://www.h2ont.com/xx.qt"
         Height          =   180
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   2340
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   330
      Left            =   4005
      TabIndex        =   1
      Top             =   720
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   2610
      TabIndex        =   0
      Top             =   720
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "         "
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   945
      Width           =   2715
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Dim filename As String
filename = Form5.Text1.Text
If Len(filename) > 0 Then
Form2.MediaPlayer1.filename = filename
Form2.Show
Unload Me
End If
End Sub
Private Sub Form_Load()
Form5.Command1.Caption = "�򿪸�ý��(&O)"
Form5.Command2.Caption = "ȡ���ò���(&Q)"
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

