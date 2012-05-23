VERSION 5.00
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "LYFTOOLS.OCX"
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "H2ont Leask Snowman Media ilxz 3.5 Plus Edition Setup Wizard"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   Icon            =   "图标h.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   7515
   StartUpPosition =   2  '屏幕中心
   Begin API控制大全.LyfTools Lf 
      Left            =   180
      Top             =   4140
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame6 
      Caption         =   "是否运行以下程序?"
      Height          =   1500
      Left            =   2115
      TabIndex        =   10
      Top             =   2205
      Width           =   5235
      Begin VB.CheckBox Check4 
         Caption         =   "我要提问!"
         Height          =   240
         Left            =   270
         TabIndex        =   15
         Top             =   1080
         Width           =   3120
      End
      Begin VB.CheckBox Check3 
         Caption         =   "访问流动网络了解 Snowman Media"
         Height          =   240
         Left            =   270
         TabIndex        =   13
         Top             =   810
         Width           =   3885
      End
      Begin VB.CheckBox Check2 
         Caption         =   "定制个性化的 Snowman Media ilxz 3.5"
         Height          =   240
         Left            =   270
         TabIndex        =   12
         Top             =   540
         Width           =   4245
      End
      Begin VB.CheckBox Check1 
         Caption         =   "立即启动 Snomwman Media ilxz 3.5"
         Height          =   240
         Left            =   270
         TabIndex        =   11
         Top             =   270
         Value           =   1  'Checked
         Width           =   4560
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "完成(F)"
      Default         =   -1  'True
      Height          =   330
      Left            =   6075
      TabIndex        =   4
      Top             =   4230
      Width           =   1185
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   7425
      TabIndex        =   3
      Top             =   3735
      Width           =   825
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   -675
      TabIndex        =   2
      Top             =   3555
      Width           =   825
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   -450
      TabIndex        =   1
      Top             =   4095
      Width           =   8295
   End
   Begin VB.Frame Frame1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   7485
   End
   Begin VB.Frame Frame5 
      Caption         =   "是否将 Snowman Media ilxz 3.5 设为默认的媒体播放器?"
      Height          =   915
      Left            =   2115
      TabIndex        =   5
      Top             =   675
      Width           =   5235
      Begin VB.OptionButton Option2 
         Caption         =   "保持计算机现有的媒体关联不变"
         Height          =   240
         Left            =   270
         TabIndex        =   9
         Top             =   540
         Width           =   4020
      End
      Begin VB.OptionButton Option1 
         Caption         =   "将 Snowman Media ilxz 3.5 设为默认的媒体播放器"
         Height          =   240
         Left            =   270
         TabIndex        =   8
         Top             =   270
         Value           =   -1  'True
         Width           =   4875
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   225
         TabIndex        =   6
         Top             =   540
         Width           =   4245
      End
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "最后:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2115
      TabIndex        =   14
      Top             =   1890
      Width           =   450
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "设置媒体关联:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2115
      TabIndex        =   7
      Top             =   360
      Width           =   1170
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   3765
      Left            =   180
      Picture         =   "图标h.frx":0442
      Top             =   180
      Width           =   1725
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim i As Integer
 Dim File As String
 Dim GB As Boolean

Private Sub Command1_Click()
On Error Resume Next
     Dim hKey As Long
     Dim MyReturn As Long
     Dim MyData As String
If Option1.Value = True Then
For i = 0 To 34
    If i = 0 Then File = ".au"
    If i = 1 Then File = ".and"
    If i = 2 Then File = ".aif"
    If i = 3 Then File = ".wmv"
    If i = 4 Then File = ".aifc"
    If i = 5 Then File = ".aiff"
    If i = 6 Then File = ".mpe"
    If i = 7 Then File = ".mpa"
    If i = 8 Then File = ".wax"
    If i = 9 Then File = ".rmi"
    If i = 10 Then File = ".asx"
    If i = 11 Then File = ".m1v"
    If i = 12 Then File = ".mp2"
    If i = 13 Then File = ".asf"
    If i = 14 Then File = ".mov"
    If i = 15 Then File = ".mp3"
    If i = 16 Then File = ".qt"
    If i = 17 Then File = ".mpeg"
    If i = 18 Then File = ".mpg"
    If i = 19 Then File = ".wma"
    If i = 20 Then File = ".wav"
    If i = 21 Then File = ".avi"
    If i = 22 Then File = ".mid"
    If i = 23 Then File = ".smi"
    If i = 24 Then File = ".smil"
    If i = 25 Then File = ".rt"
    If i = 26 Then File = ".mpv"
    If i = 27 Then File = ".rp"
    If i = 28 Then File = ".ram"
    If i = 29 Then File = ".rmm"
    If i = 30 Then File = ".rtx"
    If i = 31 Then File = ".ra"
    If i = 32 Then File = ".rm"
    If i = 33 Then File = ".m3u"
    If i = 34 Then File = ".pls"

     MyReturn = OSRegOpenKey(HKEY_CLASSES_ROOT, File, hKey)
    MyReturn = RegQueryStringValue(hKey, "", MyData)
    MyReturn = OSRegOpenKey(HKEY_CLASSES_ROOT, MyData + "\shell\open\command", hKey)
     MyReturn = RegSetStringValue(hKey, "", App.Path + "\Snowman Media ilxz 3.5.exe  %1", False)
    
   MyReturn = OSRegOpenKey(HKEY_CLASSES_ROOT, MyData + "\shell\play\command", hKey)
     MyReturn = RegSetStringValue(hKey, "", App.Path + "\Snowman Media ilxz 3.5.exe  %1", False)
        OSRegCloseKey (hKey)
 Next
End If
If Check1.Value = 1 Then Shell (App.Path + "\Snowman Media ilxz 3.5.exe")
If Check2.Value = 1 Then Shell (App.Path + "\SmM_ST.exe")
If Check3.Value = 1 Then Lf.HttpTo ("http://www.h2ont.com")
If Check4.Value = 1 Then Lf.SendMail ("leask@21cn.com")
 
 GB = True

End Sub

Private Sub Form_Load()
Me.Hide
Command1.Left = (Me.Width + Command1.Width) / 2 - Command1.Width
Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Form2 = Nothing
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If GB = True Then End
End Sub

