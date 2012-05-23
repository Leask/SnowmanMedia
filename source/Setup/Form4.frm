VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snowman Media  2.0   Setup Wizard"
   ClientHeight    =   3255
   ClientLeft      =   2880
   ClientTop       =   3180
   ClientWidth     =   6870
   Icon            =   "FORM4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6870
   Begin VB.CommandButton Command2 
      Caption         =   "不许可协议并退出安装向导(&N)"
      Height          =   330
      Left            =   3420
      TabIndex        =   9
      Top             =   2925
      Width           =   3435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "许可协议进入下一步(&Y)"
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   2925
      Width           =   3435
   End
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
      Width           =   4380
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
         Caption         =   "www.h2ont.com leask@china.com"
         Height          =   180
         Left            =   1620
         TabIndex        =   4
         Top             =   1080
         Width           =   2610
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "版本 HLS-SMM V2.2 C.TBB"
         Height          =   180
         Left            =   450
         TabIndex        =   3
         Top             =   630
         Width           =   2070
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   1500
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "FORM4.frx":1582
      Top             =   1350
      Width           =   4380
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3735
      Left            =   4410
      Picture         =   "FORM4.frx":1A8E
      ScaleHeight     =   3675
      ScaleWidth      =   6000
      TabIndex        =   0
      Top             =   -855
      Width           =   6060
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pId As Long, pHnd As Long
Const SYNCHRONIZE = &H100000
Const INFINITE = &HFFFFFFFF
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Type SHFILEOPSTRUCT
        hwnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAnyOperationsAborted As Long
        hNameMappings As Long
        lpszProgressTitle As String
End Type
Private Declare Function SHFileOperation Lib _
        "shell32" _
        (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function GetWindowsDirectory _
        Lib "kernel32" Alias "GetWindowsDirectoryA" _
        (ByVal lpBuffer As String, ByVal nSize As _
        Long) As Long
Const FO_COPY = &H2
Const FO_DELETE = &H3
Const FO_MOVE = &H1
Const FO_RENAME = &H4
Const FOF_ALLOWUNDO = &H40
Dim DirString As String
Private Sub Command1_Click()
    Dim xFile As SHFILEOPSTRUCT
    xFile.pFrom = ".\smmst\VB6CHS.DLL"
    xFile.pTo = "c:\windows\system"
    xFile.fFlags = FOF_ALLOWUNDO
    xFile.wFunc = FO_COPY
    xFile.hwnd = Me.hwnd
    If SHFileOperation(xFile) Then
    End If
   pId = Shell(".\smmst\setup.exe", vbNormalFocus)
pHnd = OpenProcess(SYNCHRONIZE, 0, pId)
If pHnd <> 0 Then
    Call WaitForSingleObject(pHnd, INFINITE)
    Call CloseHandle(pHnd)
End If
  End
  End Sub
Private Sub Command2_Click()
End
End Sub
