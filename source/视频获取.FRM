VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Video Source  － Snowman Media Pictures Browser  1.0"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   5925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraSep 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   15
      Left            =   -540
      TabIndex        =   7
      Top             =   855
      Width           =   7215
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4590
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   5
      Top             =   945
      Width           =   1095
      Begin VB.Label Label14 
         BackColor       =   &H0000FFFF&
         Caption         =   "取消(&C)"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   45
         TabIndex        =   6
         Top             =   45
         Width           =   1635
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3420
      ScaleHeight     =   255
      ScaleWidth      =   1065
      TabIndex        =   3
      Top             =   945
      Width           =   1095
      Begin VB.Label Label13 
         BackColor       =   &H0000FFFF&
         Caption         =   "确定(&Y)"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   45
         TabIndex        =   4
         Top             =   45
         Width           =   1680
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   225
      ScaleHeight     =   255
      ScaleWidth      =   5430
      TabIndex        =   1
      Top             =   45
      Width           =   5460
      Begin VB.Label lblDriver 
         BackColor       =   &H0080FFFF&
         Caption         =   "已安装的驱动程序:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   90
         TabIndex        =   2
         Top             =   45
         Width           =   4095
      End
   End
   Begin VB.ComboBox cmboSource 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   225
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   405
      Width           =   5460
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
    
    Dim lpszName As String * 100
    Dim lpszVer As String * 100
    Dim x As Integer
    Dim lResult As Long
    Dim Caps As CAPDRIVERCAPS
    
    '// Get a list of all the installed drivers
    x = 0
    Do
        lResult = capGetDriverDescriptionA(x, lpszName, 100, lpszVer, 100)   '// Retrieves driver info
        If lResult Then
            cmboSource.AddItem lpszName
            x = x + 1
        End If
    Loop Until lResult = False

    '// Get the capabilities of the current capture driver
    lResult = capDriverGetCaps(lwndC, VarPtr(Caps), Len(Caps))
    
    '// Select the driver that is currently being used
    If lResult Then cmboSource.ListIndex = Caps.wDeviceIndex

End Sub


Private Sub Label13_Click()
  
    Dim sTitle As String
    Dim Caps As CAPDRIVERCAPS
    
    If cmboSource.ListIndex <> -1 Then
        
        '// Connect the capture window to the driver
        If capDriverConnect(lwndC, cmboSource.ListIndex) Then
    
            '// Get the capabilities of the capture driver
            capDriverGetCaps lwndC, VarPtr(Caps), Len(Caps)
            
            '// If the capture driver does not support a dialog, grey it out
            '// in the menu bar.
            mfrmMain.mnuSource.Enabled = Caps.fHasDlgVideoSource
            mfrmMain.mnuFormat.Enabled = Caps.fHasDlgVideoFormat
            mfrmMain.mnuDisplay.Enabled = Caps.fHasDlgVideoDisplay
        
            sTitle = cmboSource.Text
            
            SetWindowText lwndC, sTitle
            ResizeCaptureWindow lwndC
        End If
    
    End If
    
    
    Unload Me
   
End Sub

Private Sub Label14_Click()
  Unload Me
 
End Sub
Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label13.BackColor = &HFF0000
Picture9.BackColor = &HFF0000
Label13.ForeColor = &HFFFF&
End Sub
Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label14.BackColor = &HFF0000
Picture10.BackColor = &HFF0000
Label14.ForeColor = &HFFFF&
End Sub
Private Sub form_mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)

Picture9.BackColor = &HFFFF&

Picture10.BackColor = &HFFFF&
Label13.ForeColor = &HFF0000
Label13.BackColor = &HFFFF&
Label14.ForeColor = &HFF0000
Label14.BackColor = &HFFFF&
End Sub



