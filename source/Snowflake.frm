VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择 Snowflake"
   ClientHeight    =   4995
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6645
   Icon            =   "Snowflake.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4965
      Left            =   45
      ScaleHeight     =   4965
      ScaleWidth      =   6585
      TabIndex        =   0
      Top             =   45
      Width           =   6585
      Begin VB.ListBox LF1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   4455
         IntegralHeight  =   0   'False
         ItemData        =   "Snowflake.frx":000C
         Left            =   4635
         List            =   "Snowflake.frx":000E
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         ToolTipText     =   "曲目列表"
         Top             =   45
         Width           =   1860
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "描述 :"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   90
         TabIndex        =   6
         Top             =   4590
         Width           =   6360
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "版权 :"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   90
         TabIndex        =   5
         Top             =   4320
         Width           =   4425
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "公司 :"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   90
         TabIndex        =   4
         Top             =   4050
         Width           =   4425
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "作者 :"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   90
         TabIndex        =   3
         Top             =   3780
         Width           =   4425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "标题 :"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   90
         TabIndex        =   2
         Top             =   3510
         Width           =   4425
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   2595
         Left            =   360
         Top             =   225
         Width           =   3570
      End
   End
   Begin VB.Menu dfge 
      Caption         =   "a"
      Visible         =   0   'False
      Begin VB.Menu dfgewa 
         Caption         =   "应用(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu greg 
         Caption         =   "-"
      End
      Begin VB.Menu zxcd 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu ccc 
         Caption         =   "-"
      End
      Begin VB.Menu zcee 
         Caption         =   "添加(&A)"
         Enabled         =   0   'False
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu zcsee 
         Caption         =   "更多(&M)"
         Enabled         =   0   'False
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu zee 
         Caption         =   "-"
      End
      Begin VB.Menu czEe 
         Caption         =   "退出(&X)"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long

Private Sub czEe_Click()
Unload Me
Set Form3 = Nothing
End Sub

Private Sub dfgewa_Click()
Form1.Ly.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Path", Form1.Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Path_" + Str(LF1.ListIndex))
Form1.Ly.SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Bp", Form1.Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Bp_" + Str(LF1.ListIndex))
Form1.Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_H", Form1.Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_H_" + Str(LF1.ListIndex))
Form1.Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_W", Form1.Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_W_" + Str(LF1.ListIndex))
Form1.Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vh", Form1.Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vh_" + Str(LF1.ListIndex))
Form1.Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vw", Form1.Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vw_" + Str(LF1.ListIndex))
Form1.Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vx", Form1.Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vx_" + Str(LF1.ListIndex))
Form1.Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vy", Form1.Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Vy_" + Str(LF1.ListIndex))
Form1.Ly.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Co", Str(LF1.ListIndex)
Form1.juyjk.Checked = False
Form1.juyjk_Click
Unload Me
Set Form3 = Nothing

End Sub

Private Sub dgrrevgdfg_Click()
LF1_Click

End Sub

Private Sub Form_Load()
On Error Resume Next
Form1.Ly.MakeTop Form1, False
Form1.Dir1.Path = App.Path + "\SmM_Snowflakes"
 If Form1.Dir1.ListCount = 0 Then Exit Sub
For i = 0 To Form1.Dir1.ListCount - 1
    LF1.AddItem Form1.Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Name_" + Str(i))

Next
LF1.ListIndex = Form1.Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Co")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Form1.a00308.Checked = True Then Form1.Ly.MakeTop Form1, True
Unload Me
Set Form3 = Nothing
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then PopupMenu Me.dfge, 0, Picture1.Left + Image1.Left + X, Picture1.Top + Image1.Top + Y
End Sub

Private Sub LF1_Click()
Image1.Picture = LoadPicture(Form1.Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Bi_" + Str(LF1.ListIndex)))
Label1.Caption = "标题 : " + Form1.Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Name_" + Str(LF1.ListIndex))
Label2.Caption = "作者 : " + Form1.Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Aut_" + Str(LF1.ListIndex))
Label3.Caption = "公司 : " + Form1.Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Com_" + Str(LF1.ListIndex))
Label4.Caption = "版权 : " + Form1.Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Cpy_" + Str(LF1.ListIndex))
Label5.Caption = "描述 : " + Form1.Ly.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Info_" + Str(LF1.ListIndex))
Image1.Left = Form1.Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Ix_" + Str(LF1.ListIndex))
Image1.Top = Form1.Ly.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Snowflake_Iy_" + Str(LF1.ListIndex))
End Sub

Private Sub LF1_DblClick()
dfgewa_Click
End Sub

Private Sub LF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.dfge, 0, Picture1.Left + LF1.Left + X, Picture1.Top + LF1.Top + Y
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMenu Me.dfge, 0, Picture1.Left + X, Picture1.Top + Y
End Sub
