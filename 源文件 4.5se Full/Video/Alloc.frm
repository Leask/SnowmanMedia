VERSION 5.00
Begin VB.Form frmAlloc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置文件大小"
   ClientHeight    =   2790
   ClientLeft      =   1575
   ClientTop       =   4860
   ClientWidth     =   5895
   ControlBox      =   0   'False
   Icon            =   "Alloc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "设置:"
      Height          =   1410
      Left            =   270
      TabIndex        =   3
      Top             =   630
      Width           =   5235
      Begin VB.TextBox txtAlloc 
         Height          =   330
         Left            =   1620
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "1"
         Top             =   765
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "兆字节(MBytes)"
         Height          =   285
         Left            =   2475
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "剩余磁盘空间:"
         Height          =   285
         Left            =   270
         TabIndex        =   9
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "捕获文件大小:"
         Height          =   330
         Left            =   270
         TabIndex        =   8
         Top             =   810
         Width           =   1455
      End
      Begin VB.Label lblFreeDisk 
         Caption         =   "000"
         Height          =   285
         Left            =   1665
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "兆字节(MBytes)"
         Height          =   285
         Left            =   2475
         TabIndex        =   5
         Top             =   810
         Width           =   1680
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   330
      Left            =   4410
      TabIndex        =   2
      Top             =   2205
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&Y)"
      Height          =   330
      Left            =   3060
      TabIndex        =   1
      Top             =   2205
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "情根据需要设置要捕获的文件大小,溢出的视频数据将自动删除。"
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Top             =   225
      Width           =   5235
   End
End
Attribute VB_Name = "frmAlloc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private available As Long

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Call capFileAlloc(frmMain.capwnd, txtAlloc.Text * ONE_MEGABYTE)
Unload Me
End Sub

Private Sub Form_Load()
Dim capFileSize As Long

lblFreeDisk.Caption = GetFreeSpace() - 1 'subtract 1 to compensate for possible rounding errors
capFileSize = FileLen(capFileGetCaptureFile(frmMain.capwnd))
If capFileSize > (ONE_MEGABYTE / 2) Then
    txtAlloc.Text = capFileSize / ONE_MEGABYTE
Else
    txtAlloc.Text = 1
End If
txtAlloc.SelStart = 0
txtAlloc.SelLength = Len(txtAlloc.Text)

End Sub

Private Sub txtAlloc_Change()
If Val(txtAlloc.Text) < 0 Then txtAlloc.Text = 1
If Val(lblFreeDisk.Caption) < 1 Then Exit Sub
If Val(txtAlloc.Text) > Val(lblFreeDisk.Caption) Then txtAlloc.Text = lblFreeDisk.Caption
End Sub

Private Sub txtAlloc_KeyPress(KeyAscii As Integer)
'allow backspace key
If KeyAscii = 8 Then Exit Sub
'logic to keep the user input valid
If KeyAscii < 48 Then KeyAscii = 0
If KeyAscii > 57 Then KeyAscii = 0
End Sub
