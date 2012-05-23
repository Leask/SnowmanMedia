VERSION 5.00
Begin VB.Form frmCapFrame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "捕获连续帧"
   ClientHeight    =   1875
   ClientLeft      =   2460
   ClientTop       =   3900
   ClientWidth     =   4260
   Icon            =   "CapFrame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   330
      Left            =   2880
      TabIndex        =   4
      Top             =   1305
      Width           =   1050
   End
   Begin VB.CommandButton cmdCapture 
      Caption         =   "捕获(&A)"
      Height          =   330
      Left            =   1620
      TabIndex        =   3
      Top             =   1305
      Width           =   960
   End
   Begin VB.Label lblFrames 
      Alignment       =   2  'Center
      Caption         =   "0 帧"
      Height          =   225
      Left            =   1080
      TabIndex        =   2
      Top             =   765
      Width           =   1560
   End
   Begin VB.Label lblCapFile 
      Alignment       =   2  'Center
      Height          =   225
      Left            =   405
      TabIndex        =   1
      Top             =   495
      Width           =   3480
   End
   Begin VB.Label lblPrompt 
      Caption         =   "选择源进行连续帧捕获:"
      Height          =   270
      Left            =   315
      TabIndex        =   0
      Top             =   180
      Width           =   2190
   End
End
Attribute VB_Name = "frmCapFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCapture_Click()
    If capCaptureSingleFrame(frmMain.capwnd) Then
        lblFrames.Caption = Val(lblFrames.Caption) + 1 & " 帧"
        cmdCancel.Caption = "关闭(&X)"
    Else
        MsgBox "帧捕获失败!" ', App.Title  ', vbInformation
    End If
End Sub

Private Sub Form_Load()
lblCapFile.Caption = capFileGetCaptureFile(frmMain.capwnd)
If lblCapFile.Caption = "" Then
    lblCapFile.Caption = "错误:没设置捕获文件!"
End If
Call capCaptureSingleFrameOpen(frmMain.capwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call capCaptureSingleFrameClose(frmMain.capwnd)
End Sub
