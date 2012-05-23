VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Snowman Media  Playing Lists Editer"
   ClientHeight    =   1035
   ClientLeft      =   2760
   ClientTop       =   3690
   ClientWidth     =   4530
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "取消(&Q)"
      Height          =   375
      Left            =   2250
      TabIndex        =   3
      Top             =   630
      Width           =   2265
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定(&Y)"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   630
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      Caption         =   "请输入新建列表的名称(&N):"
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4515
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   4335
      End
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim filename As String
Dim filenum As Integer
Private Sub Command1_Click()
filename = Form6.Dir1.Path & "\" & Dialog.Text1.Text & ".m3u"
Form6.Frame2.Caption = filename
filenum = FreeFile
Open filename For Output As #filenum
Form6.File1.Path = Form6.Dir1.Path
Form6.Text1.Text = "# 列表已成功创建,请选择喜欢的媒体内容加入列表(包括网络媒体)。"
Form6.Text2.Text = Dialog.Text1.Text & ".m3u"
Close #filenum
Form6.File1.Path = "c:"
Form6.File1.Path = Form6.Dir1.Path
Unload Me
End Sub
Private Sub command2_Click()
Unload Me
End Sub

