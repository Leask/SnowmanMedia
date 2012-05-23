VERSION 5.00
Begin VB.Form frmCapPal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������ɫ��"
   ClientHeight    =   1365
   ClientLeft      =   2205
   ClientTop       =   1890
   ClientWidth     =   5145
   Icon            =   "CapPal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtColors 
      Height          =   285
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "256"
      Top             =   270
      Width           =   465
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   330
      Left            =   3825
      TabIndex        =   2
      Top             =   765
      Width           =   1005
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "��ʼ(&S)"
      Height          =   330
      Left            =   2565
      TabIndex        =   1
      Top             =   765
      Width           =   1005
   End
   Begin VB.CommandButton cmdFrame 
      Caption         =   "֡(&F)"
      Height          =   330
      Left            =   1305
      TabIndex        =   0
      Top             =   765
      Width           =   1005
   End
   Begin VB.Label lblFrames 
      Alignment       =   2  'Center
      Caption         =   "0 ֡"
      Height          =   240
      Left            =   1935
      TabIndex        =   4
      Top             =   315
      Width           =   2805
   End
   Begin VB.Label lblColors 
      Caption         =   "��ɫ:"
      Height          =   240
      Left            =   315
      TabIndex        =   3
      Top             =   315
      Width           =   645
   End
End
Attribute VB_Name = "frmCapPal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'form level flag to indicate whether
'we need to close the palette capture on unload
Private fManual As Boolean
'form level flag to record number of frames captured in manual mode
Private numManFrames As Long

Private Sub Form_Load()
'load num pal colors from registry
    txtColors.Text = GetSetting(App.Title, "��ɫ��", "��ɫ��", "256")
End Sub

Private Sub cmdFrame_Click()
    fManual = True
    Call capPaletteManual(frmMain.capwnd, False, Val(txtColors.Text))
    numManFrames = numManFrames + 1
    lblFrames.Caption = numManFrames & " ֡"
    cmdCancel.Caption = "�ر�(&X)"
End Sub

Private Sub cmdStart_Click()
    Const numFrames As Long = 100 'change this value to sample more or less frames
    numManFrames = 0 'reset manual capture counter if necessary
    fManual = False
    lblFrames.Caption = "ȡ����,���Ժ�..."
    lblFrames.Refresh
    cmdFrame.Enabled = False
    Call capPaletteAuto(frmMain.capwnd, numFrames, Val(txtColors.Text))
    lblFrames.Caption = "ʧ��! - " & numFrames & " ֡��ȡ!"
    cmdFrame.Enabled = True
    cmdCancel.Caption = "ȷ��(&O)"
End Sub

Private Sub txtColors_KeyPress(KeyAscii As Integer)
    'allow backspace key
    If KeyAscii = 8 Then Exit Sub
    'logic to keep the user input valid
    If KeyAscii < 48 Then KeyAscii = 0
    If KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtColors_LostFocus()
    'Input Filter
    If Val(txtColors.Text) < 16 Then txtColors.Text = 16
    If Val(txtColors.Text) > 256 Then txtColors.Text = 256
End Sub

Private Sub cmdCancel_Click()
    If fManual Then
        'close manual palette capture by sending false
        Call capPaletteManual(frmMain.capwnd, False, Val(txtColors.Text))
    End If
    If cmdCancel.Caption <> "ȡ��(&C)" Then 'save num colors to registry
        Call SaveSetting(App.Title, "��ɫ��", "��ɫ��", txtColors.Text)
    End If
    Unload Me
End Sub
