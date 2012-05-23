VERSION 5.00
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "LYFTOOLS.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Sm.M. Write Reg"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "setupreg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.Frame Frame16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   2745
      TabIndex        =   32
      Top             =   1260
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   13
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   4005
      TabIndex        =   31
      Top             =   1305
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   15
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   5040
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   6300
      TabIndex        =   29
      Top             =   1125
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   12
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   2835
      TabIndex        =   28
      Top             =   405
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   11
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   4095
      TabIndex        =   27
      Top             =   450
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   14
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   5130
      TabIndex        =   26
      Top             =   225
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   6390
      TabIndex        =   25
      Top             =   270
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   16
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   5040
      TabIndex        =   24
      Top             =   2835
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   6300
      TabIndex        =   23
      Top             =   2880
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   2430
      TabIndex        =   22
      Top             =   2880
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   3690
      TabIndex        =   21
      Top             =   2925
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   5130
      TabIndex        =   20
      Top             =   1980
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   17
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   1545
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   6390
      TabIndex        =   19
      Top             =   2025
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   10
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   2520
      TabIndex        =   18
      Top             =   2025
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   3780
      TabIndex        =   17
      Top             =   2070
      Visible         =   0   'False
      Width           =   1050
      Begin VB.OptionButton Option 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   -45
         TabIndex        =   34
         Top             =   90
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.CheckBox Check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   14
      Left            =   1800
      TabIndex        =   16
      Top             =   8025
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CheckBox Check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   13
      Left            =   1845
      TabIndex        =   15
      Top             =   7080
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CheckBox Check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   12
      Left            =   1845
      TabIndex        =   14
      Top             =   7530
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CheckBox Check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   11
      Left            =   315
      TabIndex        =   13
      Top             =   8010
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CheckBox Check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   10
      Left            =   405
      TabIndex        =   12
      Top             =   7560
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CheckBox Check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   9
      Left            =   90
      TabIndex        =   11
      Top             =   7155
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CheckBox Check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   8
      Left            =   360
      TabIndex        =   10
      Top             =   6750
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CheckBox Check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   7
      Left            =   2250
      TabIndex        =   9
      Top             =   6210
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CheckBox Check 
      Caption         =   "Check1"
      Height          =   330
      Index           =   6
      Left            =   1613
      TabIndex        =   8
      Top             =   -390
      Width           =   1410
   End
   Begin VB.CheckBox Check 
      Caption         =   "Check1"
      Height          =   330
      Index           =   5
      Left            =   1568
      TabIndex        =   7
      Top             =   -1110
      Width           =   1410
   End
   Begin VB.CheckBox Check 
      Caption         =   "Check1"
      Height          =   330
      Index           =   4
      Left            =   1568
      TabIndex        =   6
      Top             =   -660
      Width           =   1410
   End
   Begin VB.CheckBox Check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   3
      Left            =   315
      TabIndex        =   5
      Top             =   6255
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CheckBox Check 
      Caption         =   "Check1"
      Height          =   330
      Index           =   2
      Left            =   3053
      TabIndex        =   4
      Top             =   -345
      Width           =   1410
   End
   Begin VB.CheckBox Check 
      Caption         =   "Check1"
      Height          =   330
      Index           =   1
      Left            =   3098
      TabIndex        =   3
      Top             =   -795
      Width           =   1410
   End
   Begin VB.CheckBox Check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   2655
      TabIndex        =   2
      Top             =   6840
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   4050
      TabIndex        =   1
      Top             =   7875
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox Text 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1185
      Index           =   1
      Left            =   4275
      TabIndex        =   0
      Top             =   6660
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9090
      Top             =   5985
   End
   Begin API控制大全.LyfTools Lyf 
      Left            =   7380
      Top             =   5850
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim UnDt As Boolean
Private Sub Form_Load()
UnDt = False

Check(14).Value = 1
Check(0).Value = 1
Check(1).Value = 1
Check(2).Value = 0
Check(4).Value = 0
Check(5).Value = 0
Check(12).Value = 0
Check(6).Value = 1
Check(7).Value = 1
Check(8).Value = 0
Check(9).Value = 1
Check(10).Value = 1
Check(11).Value = 1
Check(3).Value = 0
Check(13).Value = 1

Me.Option(14).Value = 1
Me.Option(15).Value = 0
Me.Option(16).Value = 0
Me.Option(17).Value = 0
Me.Option(2).Value = 1
Me.Option(4).Value = 0
Me.Option(3).Value = 1
Me.Option(5).Value = 0
Me.Option(7).Value = 0
Me.Option(6).Value = 1
Me.Option(8).Value = 0
Me.Option(9).Value = 0
Me.Option(10).Value = 0
Me.Option(11).Value = 1
Me.Option(12).Value = 1
Me.Option(13).Value = 0
Text(2).Text = App.Path + "\BPics\BCG_115.gif"
Text(1).Text = "163"

For I = 2 To 17
Lyf.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Op_" + Str(I), Me.Option(I).Value
Next
For I = 0 To 14
Lyf.SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Ch_" + Str(I), Me.Check(I).Value
Next
For I = 1 To 2
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Te_" + Str(I), Me.Text(I).Text
Next













Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Soft_Path", App.Path

Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_ 1_Path", App.Path + "\Sflakes\Sm.M.3.0"
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_ 2_Path", App.Path + "\Sflakes\蛋蛋"
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_ 3_Path", App.Path + "\Sflakes\Jurassic"
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_ 1_Name", "Sm.M.3.0"
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_ 2_Name", "蛋蛋"
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_ 3_Name", "Jurassic"

Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_ 1_Path", App.Path + "\Skins\Media"
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_ 2_Path", App.Path + "\Skins\Neoni"
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_ 3_Path", App.Path + "\Skins\Yellow"
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_ 1_Name", "Media"
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_ 2_Name", "Neoni"
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_ 3_Name", "Yellow"


Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path", App.Path + "\Skins\Media"
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Sflake_Path", App.Path + "\Sflakes\Sm.M.3.0"
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Te_ 2", App.Path + "\Bpics\BCG_104.gif"
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "VolID", "3.06.0736 Lxis Basic"
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "VolDay", "10.1.2001"
Lyf.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "RunTime", "One"

UnDt = True



End Sub

Private Sub Timer1_Timer()
If UnDt = True Then End
End Sub
