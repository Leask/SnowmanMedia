VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{822CBFD1-3FFF-11D7-ACD6-0050BAC05F28}#9.0#0"; "CURTMENU.OCX"
Begin VB.Form frmtest 
   AutoRedraw      =   -1  'True
   Caption         =   "Test form"
   ClientHeight    =   2925
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0080FFFF&
   Icon            =   "frmtest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   4470
   StartUpPosition =   1  '所有者中心
   Begin CurtMenu嵌入式图形菜单.CurtMenu CurtMenu1 
      Left            =   1170
      Top             =   2160
      _ExtentX        =   1588
      _ExtentY        =   741
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3330
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":030A
            Key             =   ""
            Object.Tag             =   "新建"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":0466
            Key             =   ""
            Object.Tag             =   "选择"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":05C2
            Key             =   ""
            Object.Tag             =   "打开"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":071E
            Key             =   ""
            Object.Tag             =   "保存"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":087A
            Key             =   ""
            Object.Tag             =   "复制"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":09D6
            Key             =   ""
            Object.Tag             =   "剪切"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":0B32
            Key             =   ""
            Object.Tag             =   "粘贴"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":0C8E
            Key             =   ""
            Object.Tag             =   "退出"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuNew 
         Caption         =   "新建"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "打开"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "保存"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveas 
         Caption         =   "另存为(&A)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "退出"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuCut 
         Caption         =   "剪切"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "复制"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "粘贴"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuOptions 
         Caption         =   "选择(&O)"
         Begin VB.Menu mnuSubA 
            Caption         =   "子菜单"
            Index           =   0
         End
         Begin VB.Menu mnuSubA 
            Caption         =   "子菜单"
            Checked         =   -1  'True
            Index           =   1
         End
      End
      Begin VB.Menu mnuSub 
         Caption         =   "效果A"
         Index           =   0
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuSub 
         Caption         =   "效果B"
         Index           =   1
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuSub 
         Caption         =   "效果C"
         Index           =   2
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuSub 
         Caption         =   "效果D"
         Checked         =   -1  'True
         Index           =   3
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuSub 
         Caption         =   "效果E"
         Index           =   4
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuWWW 
         Caption         =   "查看最新版本"
      End
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'对菜单进行转换
Private Sub Form_Load()
    CurtMenu1.Connect Me.hWnd, True, ImageList1
End Sub

'弹出菜单的演示
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuEdit
    End If
End Sub

'菜单Enabled属性的演示
Private Sub mnuCopy_Click()
  mnuPaste.Enabled = True
  mnuCopy.Enabled = False
  mnuCut.Enabled = False
End Sub
Private Sub mnuCut_Click()
  mnuPaste.Enabled = True
  mnuCopy.Enabled = False
  mnuCut.Enabled = False
End Sub
Private Sub mnuPaste_Click()
  mnuPaste.Enabled = False
  mnuCopy.Enabled = True
  mnuCut.Enabled = True
End Sub

'菜单Checked属性的演示
Private Sub mnuSubA_Click(Index As Integer)
    mnuSubA(Index).Checked = Not mnuSubA(Index).Checked
End Sub


'菜单Click事件的演示
Private Sub mnuQuit_Click()
  Unload Me
End Sub
Private Sub mnuSub_Click(Index As Integer)
Dim I As Long
    For I = 0 To mnuSub.Count - 1
        mnuSub(I).Checked = False
    Next
    mnuSub(Index).Checked = 1
    CurtMenu1.HoverFillStyle = Index
End Sub
'访问网站获取更多最新CurSoft产品
Private Sub mnuWWW_Click()
    ShellExecute 0&, "Open", "http://www.curtsoft.com", "", App.Path, 1
End Sub

'********************************用户须知************************************************************
'
'   CurtMenu v1.01  有条件免费使用！
'   版权归所有:CurtSoft 保留一切权利!     http://www.curtsoft.com
'
'   我们只想把时间用于更多更好作品的开发，请使用者自觉遵守下面的用户协议：
'   如果您未将本控件用与商业目的，可以免费使用本控件！否则请付费：个人用户￥39，单位用户￥129。
'
'   注册方法：将注册费和注册信息发给我，收到注册协议后即注册成功。
'   联系方式：  Email：Inthenet@163.net      Mobile：13670102745     QQ：121728839
'   开户行：深圳招商银行振华路分行 帐号：0755-36387681
'   地址：深圳市福田区振华路78号电子器材大厦东418  邮编：518031
'   欢迎您对本控件提出宝贵意见，我将认真改正！谢谢使用！
'   注意：您发出注册费后请用电子邮件告诉我您的注册信息（用户名、证件号码、联系地址）！
'
'*******************************功能概要***************************************************************
'
'本控件特点如下：
'    1-使用极为简单：用一个函数将效果嵌入VB自带的菜单，并不改变其使用方法；
'    2-外观控制简单灵活，多种填充效果惊艳；
'    3-兼容WIN9X、WINNT，WIN2K，WINXP；
'
'******************************最新更新记录*****************************************************************
'2002-02-26：升级为v1.0.1：扩展两种填充效果，改进算法，提高填充速度。
'
'2002-02-25：控件发布，版本v1.0.0
'
'******************************使用说明****************************************************************
'=================方法==================
'
'改变/恢复窗体中菜单的显示效果:
'Connect(hWnd As Long, Flag As Boolean, Optional imlMenu As Object)
'    hWnd：包含要进行转换的菜单的窗体的句柄；
'    Flag：True为进行转换，False为解除转换；
'    imlMenu：包含菜单要使用图标的ImageList控件（解除转换可省略该参数）。
'注释：要将菜单ITEM和ICON进行关联，须设置菜单项的标题（Caption）和ImageList控件中ICON的标记（Tag）相同。
'==================属性=================
'BackColor返回/设置菜单条的背景颜色
'ForeColor返回/设置菜单文字的颜色
'IconBarColor返回/设置图标栏的背景颜色
'TextBarColor返回/设置文字栏的背景颜色
'ShadowColor返回/设置阴影颜色
'DisabledColor返回/设置菜单项被禁用的颜色
'CheckMarkColor返回/设置菜单项被复选时的颜色
'HoverFillStyle返回/设置菜单项高亮显示的填充类型
'HoverEdgeColor返回/设置菜单项高亮显示的边框颜色
'HoverForeColor菜单项高亮显示的文字颜色
'SepraterColor返回/设置分隔线的颜色
'
'******************************谢谢您阅读本文件***********************************************
