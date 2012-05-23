VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form plm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Playlist Maker"
   ClientHeight    =   3645
   ClientLeft      =   3540
   ClientTop       =   330
   ClientWidth     =   7725
   Icon            =   "plm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   182.25
   ScaleMode       =   2  'Point
   ScaleWidth      =   386.25
   Begin VB.CommandButton DEL 
      Caption         =   "清除(&S)"
      Height          =   375
      Left            =   2655
      TabIndex        =   7
      Top             =   3240
      Width           =   2490
   End
   Begin VB.ListBox List_v 
      Height          =   2760
      Left            =   5175
      TabIndex        =   10
      ToolTipText     =   "Click to delete a file"
      Top             =   45
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "过滤(&R)"
      Height          =   375
      Left            =   90
      TabIndex        =   9
      Top             =   3240
      Width           =   2490
   End
   Begin VB.DirListBox Dir1 
      Height          =   2400
      Left            =   45
      TabIndex        =   8
      Top             =   405
      Width           =   2535
   End
   Begin VB.ListBox List 
      Height          =   240
      Left            =   5220
      TabIndex        =   6
      ToolTipText     =   "Click to delete a file"
      Top             =   2925
      Width           =   2475
   End
   Begin VB.CommandButton Command2 
      Caption         =   "生成播放列表(&M)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5220
      TabIndex        =   5
      Top             =   3240
      Width           =   2490
   End
   Begin VB.CommandButton all 
      Caption         =   "全选(&A)"
      Height          =   375
      Left            =   2655
      TabIndex        =   4
      Top             =   2880
      Width           =   2490
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "plm.frx":1582
      Left            =   1170
      List            =   "plm.frx":1584
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Text            =   "*.*"
      Top             =   2880
      Width           =   1440
   End
   Begin VB.FileListBox File1 
      Height          =   2790
      Left            =   2610
      OLEDragMode     =   1  'Automatic
      System          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Click to select a file"
      Top             =   45
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog save 
      Left            =   990
      Top             =   4995
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "Play List"
      Filter          =   "Play List File(*.M3U)|*.M3U"
   End
   Begin VB.Label Label1 
      Caption         =   "文件类型(&T): "
      Height          =   360
      Left            =   45
      TabIndex        =   2
      Top             =   2925
      Width           =   1905
   End
End
Attribute VB_Name = "plm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub all_Click()
For i = 0 To File1.ListCount - 1
  List.AddItem Dir1.Path + "\" + File1.List(i)
  List_v.AddItem File1.List(i)
Next i
End Sub
Private Sub Command1_Click()
File1.Pattern = Combo1.Text
End Sub
Private Sub Command2_Click()
 If List.ListCount = 0 Then
  MsgBox ("NO files selected!"), vbCritical
  Exit Sub
 End If
 save.ShowSave
   Open save.filename For Output As #1
    For i = 0 To List.ListCount - 1
     Print #1, List.List(i)
    Next i
   Close (1)
End Sub
Private Sub DEL_Click()
 List.Clear
 List_v.Clear
End Sub
Private Sub Dir1_Change()
File1.Pattern = Combo1.Text
File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
Private Sub File1_Click()
 List.AddItem Dir1.Path + "\" + File1.filename
 List_v.AddItem File1.filename
End Sub
Private Sub Form_Load()
Combo1.AddItem "*.*"
Combo1.AddItem "*.MP3"
Combo1.AddItem "*.MID"
Combo1.AddItem "*.AVI"
Combo1.AddItem "*.MPG"
Combo1.AddItem "*.MOV"
Combo1.AddItem "*.DAT"
File1.Pattern = Combo1.Text
End Sub
Private Sub List_Click()
List.RemoveItem List.ListIndex
'List_v.RemoveItem List.ListIndex
Dir1.Path = Drive1.Drive
File1.Pattern = Combo1.Text
File1.Path = Dir1.Path
End Sub
Private Sub List_v_Click()
List_v.RemoveItem List_v.ListIndex
End Sub
