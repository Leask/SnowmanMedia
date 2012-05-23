VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Snowman Media  Play Lists"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5640
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   645
      Left            =   360
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   4995
      Width           =   3300
   End
   Begin VB.Frame Frame2 
      Caption         =   "没有可显示的列表内容:"
      Height          =   1590
      Left            =   0
      TabIndex        =   10
      Top             =   2835
      Width           =   5640
      Begin VB.TextBox Text1 
         Height          =   1275
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Text            =   "Form6.frx":0442
         ToolTipText     =   "你可以在这里手工编辑你的媒体播放列表  播放状态提示栏  －－Snowman Media"
         Top             =   225
         Width           =   5550
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出程序(&Q)"
      Height          =   375
      Left            =   2970
      TabIndex        =   9
      Top             =   2340
      Width           =   2625
   End
   Begin VB.CommandButton Command8 
      Caption         =   "重置所选列表(&R)"
      Height          =   375
      Left            =   1395
      TabIndex        =   8
      Top             =   2340
      Width           =   1590
   End
   Begin VB.CommandButton Command6 
      Caption         =   "把选中的媒体添加到列表(&A)"
      Height          =   375
      Left            =   2970
      TabIndex        =   7
      Top             =   1980
      Width           =   2625
   End
   Begin VB.CommandButton Command7 
      Caption         =   "修改该列表(&L)"
      Height          =   375
      Left            =   45
      TabIndex        =   6
      Top             =   2340
      Width           =   1365
   End
   Begin VB.CommandButton Command3 
      Caption         =   "新建列表(&N)"
      Height          =   375
      Left            =   1395
      TabIndex        =   5
      Top             =   1980
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      Caption         =   "播放该文件(&O)"
      Height          =   375
      Left            =   45
      TabIndex        =   4
      Top             =   1980
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Caption         =   "Snowman Media  Play Lists Editer"
      Height          =   1050
      Left            =   0
      TabIndex        =   3
      Top             =   1755
      Width           =   5640
   End
   Begin VB.FileListBox File1 
      Height          =   1710
      Left            =   2970
      TabIndex        =   2
      Top             =   45
      Width           =   2670
   End
   Begin VB.DirListBox Dir1 
      Height          =   1350
      Left            =   0
      TabIndex        =   1
      Top             =   405
      Width           =   2940
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   2940
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As String
Dim num As Integer
Dim f As String
Dim filename As String
Dim filenum As Integer
Dim i As Integer



Private Sub command2_Click()
Unload Me
End Sub

Private Sub Command7_Click()
Form11.Show

'num = File1.ListCount
'filenum = FreeFile
'Open "C:\filelist.m3u" For Output As #filenum
'For i = 0 To num - 1
'If File1.Selected(i) Then
'filename = File1.Path + "\" + File1.List(i)
'End If
'Print #filenum, filename
'Next
'Close #filenum

End Sub

Private Sub Command8_Click()
Form11.Show
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1
End Sub
Private Sub Command1_Click()
num = File1.ListCount
filenum = FreeFile
Open "C:\filelist.m3u" For Output As #filenum
For i = 0 To num - 1
If File1.Selected(i) Then
filename = File1.Path + "\" + File1.List(i)
End If
Print #filenum, filename
Next
Close #filenum
Form1.MediaPlayer1.filename = "c:\filelist.m3u"
End Sub

Private Sub command3_click()
Dialog.Show
End Sub
Private Sub command6_click()
f = Form6.Text2.Text
num = File1.ListCount
filenum = FreeFile
Open f For Append As #filenum
For i = 0 To num - 1
If File1.Selected(i) Then
filename = File1.Path + "\" + File1.List(i)
End If
Print #filenum, filename
Next
Close #filenum
Form6.Text1.Text = "# 曲目已成功加入列表,你可以重复以上步骤加入曲目或者退出以结束编辑(列表将自动保存)。"
End Sub

