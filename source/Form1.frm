VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Flash on Line Opening Window - Snowman Media  2.0"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5430
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   330
      Left            =   4005
      TabIndex        =   4
      Top             =   720
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   2610
      TabIndex        =   5
      Top             =   720
      Width           =   1410
   End
   Begin VB.Frame Frame2 
      Caption         =   "例如下列地址:"
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   630
      Width           =   2535
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "http://www.h2ont.com/x.swf"
         Height          =   180
         Left            =   90
         TabIndex        =   3
         Top             =   180
         Width           =   2340
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "网络Flash地址:"
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   5235
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Dim filename As String
filename = Form1.Text1.Text
If Len(filename) > 0 Then
        Form7.Show
        Form7.ShockwaveFlash1.Movie = filename
  Unload Me
End If
End Sub
Private Sub Form_Load()
Form1.Command1.Caption = "打开该Flash(&O)"
Form1.Command2.Caption = "取消该操作(&Q)"
End Sub
Private Sub Command2_Click()
Unload Me
End Sub


