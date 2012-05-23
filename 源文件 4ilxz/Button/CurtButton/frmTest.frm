VERSION 5.00
Object = "{69958DD9-23E5-11D6-ACD7-0050BAC05F28}#11.0#0"; "CurtButton.ocx"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   4365
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   240
      ScaleHeight     =   825
      ScaleWidth      =   3885
      TabIndex        =   8
      Top             =   840
      Width           =   3885
      Begin CurtButton多风格按钮控件.CurtButton CurtButton2 
         Height          =   735
         Index           =   0
         Left            =   2790
         TabIndex        =   9
         Top             =   90
         Width           =   915
         _extentx        =   1614
         _extenty        =   1296
         picture         =   "frmTest.frx":0000
         font            =   "frmTest.frx":001E
      End
      Begin CurtButton多风格按钮控件.CurtButton CurtButton2 
         Height          =   735
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   90
         Width           =   915
         _extentx        =   1614
         _extenty        =   1296
         picture         =   "frmTest.frx":0042
         font            =   "frmTest.frx":0060
      End
      Begin CurtButton多风格按钮控件.CurtButton CurtButton2 
         Height          =   735
         Index           =   2
         Left            =   1170
         TabIndex        =   11
         Top             =   90
         Width           =   915
         _extentx        =   1614
         _extenty        =   1296
         picture         =   "frmTest.frx":0084
         font            =   "frmTest.frx":00A2
      End
      Begin CurtButton多风格按钮控件.CurtButton CurtButton2 
         Height          =   735
         Index           =   3
         Left            =   2070
         TabIndex        =   12
         Top             =   0
         Width           =   915
         _extentx        =   1614
         _extenty        =   1296
         picture         =   "frmTest.frx":00C6
         font            =   "frmTest.frx":00E4
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   90
      ScaleHeight     =   465
      ScaleWidth      =   4515
      TabIndex        =   2
      Top             =   180
      Width           =   4515
      Begin CurtButton多风格按钮控件.CurtButton CurtButton1 
         Height          =   375
         Index           =   0
         Left            =   630
         TabIndex        =   3
         Top             =   180
         Width           =   825
         _extentx        =   1455
         _extenty        =   661
         picture         =   "frmTest.frx":0108
         font            =   "frmTest.frx":0126
      End
      Begin CurtButton多风格按钮控件.CurtButton CurtButton1 
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   825
         _extentx        =   1455
         _extenty        =   661
         picture         =   "frmTest.frx":014A
         font            =   "frmTest.frx":0168
      End
      Begin CurtButton多风格按钮控件.CurtButton CurtButton1 
         Height          =   375
         Index           =   2
         Left            =   2610
         TabIndex        =   5
         Top             =   0
         Width           =   825
         _extentx        =   1455
         _extenty        =   661
         picture         =   "frmTest.frx":018C
         font            =   "frmTest.frx":01AA
      End
      Begin CurtButton多风格按钮控件.CurtButton CurtButton1 
         Height          =   375
         Index           =   4
         Left            =   3330
         TabIndex        =   6
         Top             =   180
         Width           =   825
         _extentx        =   1455
         _extenty        =   661
         picture         =   "frmTest.frx":01CE
         font            =   "frmTest.frx":01EC
      End
      Begin CurtButton多风格按钮控件.CurtButton CurtButton1 
         Height          =   375
         Index           =   3
         Left            =   1620
         TabIndex        =   7
         Top             =   90
         Width           =   825
         _extentx        =   1455
         _extenty        =   661
         picture         =   "frmTest.frx":0210
         font            =   "frmTest.frx":022E
      End
   End
   Begin CurtButton多风格按钮控件.CurtButton CurtButton4 
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   2280
      Width           =   915
      _extentx        =   1614
      _extenty        =   661
      picture         =   "frmTest.frx":0252
      font            =   "frmTest.frx":0270
   End
   Begin CurtButton多风格按钮控件.CurtButton CurtButton3 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2280
      Width           =   915
      _extentx        =   1614
      _extenty        =   661
      picture         =   "frmTest.frx":0294
      font            =   "frmTest.frx":02B2
   End
   Begin CurtButton多风格按钮控件.CurtButton CurtButton5 
      Height          =   465
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   3255
      _extentx        =   5741
      _extenty        =   820
      picture         =   "frmTest.frx":02D6
      font            =   "frmTest.frx":02F4
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
Dim i As Integer
    Picture1.ScaleMode = vbPixels
    Picture1.BorderStyle = 0
    Picture2.ScaleMode = vbPixels
    Picture2.BorderStyle = 0

    'bt3D型的按钮，可做OFFICE97类型的工具栏
    For i = 0 To 4
        CurtButton1(i).Appearance = bt3D
        CurtButton1(i).Move 4 + i * 24, 2, 24, 24
        CurtButton1(i).ToolTipText = "bt3D"
        Set CurtButton1(i).Picture = LoadPicture(App.Path & "\" & CStr(i + 1) & ".ico")
    Next
    'btXP型的按钮，可做XP型的工具栏
    For i = 0 To 3
        CurtButton2(i).Appearance = btXP
        CurtButton2(i).Move 4 + i * 58, 2, 58, 50
        CurtButton2(i).ToolTipText = "btXP"
        Set CurtButton2(i).Picture = LoadPicture(App.Path & "\" & CStr(i + 6) & ".ico")
    Next
    'btXPplus型的按钮，可用于按钮
    With CurtButton3
        .Appearance = btXPplus
        .ShowFocus = True
        .Default = True
        .ToolTipText = "btXPplus"
        .HoverFillStyle = hsColumn
        .Caption = "确定"
    End With
    With CurtButton4
        .Appearance = btXPplus
        .ShowFocus = True
        .Cancel = True
        .ToolTipText = "btXPplus"
        .HoverFillStyle = hsLtlByLtl
        .Caption = "取消"
    End With
    'btLabel型的按钮，可用于标签
    With CurtButton5
        .Appearance = btLabel
        .Alignment = alnCenterMiddle
        .ToolTipText = "btLabel"
        .Caption = "http://www.curtsoft.com"
    End With
End Sub
'本控件给您提供了MouseEnter和MouseLeave事件，将极大方便您编程
Private Sub CurtButton5_MouseEnter()
    CurtButton5.Font.Bold = True
    CurtButton5.ForeColor = vbBlue
End Sub
Private Sub CurtButton5_MouseLeave()
    CurtButton5.Font.Bold = False
    CurtButton5.ForeColor = vbBlack
End Sub
Private Sub CurtButton5_Click()
    ShellExecute 0&, "Open", "http://www.curtsoft.com", "", App.Path, 1
End Sub

'演示快捷键和DEFAULT和CANCEL属性
Private Sub CurtButton3_Click()
    MsgBox "谢谢您试用本控件！"
End Sub

Private Sub CurtButton4_Click()
    Unload Me
End Sub


Private Sub Form_Resize()
    Picture1.Move 0, 0, Me.ScaleWidth, 420
    Picture1.Refresh
    Picture2.Move 0, Picture1.Height + 2, Me.ScaleWidth, 815
    Picture2.Refresh
End Sub
Private Sub Picture1_Paint()
    Picture1.Cls
    Picture1.Line (0, 0)-(Picture1.ScaleWidth - 1, 0), vbWhite
    Picture1.Line (0, 0)-(0, Picture1.ScaleHeight - 1), vbWhite
    Picture1.Line (Picture1.ScaleWidth - 1, 0)-(Picture1.ScaleWidth - 1, Picture1.ScaleHeight), RGB(64, 64, 64)
    Picture1.Line (0, Picture1.ScaleHeight - 1)-(Picture1.ScaleWidth, Picture1.ScaleHeight - 1), RGB(64, 64, 64)
End Sub
Private Sub Picture2_Paint()
    Picture2.Cls
    Picture2.Line (0, 0)-(Picture2.ScaleWidth - 1, 0), vbWhite
    Picture2.Line (0, 0)-(0, Picture2.ScaleHeight - 1), vbWhite
    Picture2.Line (Picture2.ScaleWidth - 1, 0)-(Picture2.ScaleWidth - 1, Picture2.ScaleHeight), RGB(64, 64, 64)
    Picture2.Line (0, Picture2.ScaleHeight - 1)-(Picture2.ScaleWidth, Picture2.ScaleHeight - 1), RGB(64, 64, 64)
End Sub
'********************************用户须知************************************************************
'
'   CurtButton v1.04  有条件免费使用！
'   版权归所有:CurtSoft 保留一切权利!     http://www.curtsoft.com
'
'   我们只想把时间用于更多更好作品的开发，请使用者自觉遵守下面的用户协议：
'   如果您未将本控件用与商业目的，可以免费使用本控件！否则请付费：个人用户￥29，单位用户￥99。
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
'    1-拥有XPplus（适合做按钮）、XP、3D（适合做工具栏）、Label（可做标签）四种风格。
'    2-提供完美的事件过程(MouseEnter,MouseLeave等），极大方便编程(见事件说明）。
'    3-XP、XPplus风格的阴影效果生动活泼，通过几个属性就可灵活控制按钮外观。
'
'******************************最新更新记录*****************************************************************
'
'本控件特点如下：
'    1-拥有多种风格，可用于按钮，工具拦，标签等，用途广泛；
'    2-提供完美的事件过程，例如MouseLeave，极大方便编程；
'    3-显示效果生动活泼，外观控制简单灵活；
'
'******************************最新更新记录*****************************************************************
'2002-02-25：修正了填充效果设置为hsLtlByLtl时应用程序无法正常显示的问题；
'
'2002-02-17:增加HoverFillStyle属性，按钮效果更加丰富；
'           修正了点击按钮弹出对话框后按钮无法复原的BUG。
'
'2002-01-28：控件发布，版本v1.0.0
'
'******************************使用说明****************************************************************
'==================属性=================
'
'所有风格按钮都使用的属性：
'Appearance--返回/设置按钮风格
'Caption--返回/显示按钮文本
'Picture--返回/显示按钮显示的图片（btLabel风格将忽略该属性）
'BackColor--返回/设置按钮的背景颜色
'ForeColor--返回/设置按钮的前景颜色
'Font--返回/设置显示文本使用的字体
'Enabled--返回/设置按钮是否可用
'MouseIcon--返回/设置按钮的自定义鼠标
'MousePointer--返回/设置按钮的系统鼠标
'Cancel--返回/设置按钮是否为窗体的“取消”按钮
'Default--返回/设置按钮是否为窗体的缺省按钮
'
'XP和XPplus风格按钮使用的属性：
'HoverFillStyle返回/设置鼠标在按钮上时的填充类型
'HoverColor--返回/显示鼠标移动到按钮上的填充颜色
'MouseDownColor--返回/设置鼠标键按下时的填充颜色
'EdgeColor--返回/显示鼠标移动到按钮上的边框颜色
'ShadowOffSet--返回/设置图标和阴影的位移量
'ShadowColor--返回/设置阴影颜色
'
'XPplus风格按钮使用的属性：
'BorderColor返回/显示鼠标移出到按钮时的边框颜色
'ShowFocus返回/设置按钮是否在获得焦点时显示焦点
'
'btLabel风格按钮使用的属性：
'Alignment返回/设置文本的对齐方式
'
'=================事件==================
'MouseEnter--鼠标进入按钮时发生
'MouseLeave--鼠标离开按钮时发生
'MouseDown--鼠标任意键在按钮上按下都发生该事件，如鼠标未UP或离开按钮，该事件将持续发生
'MouseUp--鼠标任意键在按钮上按下并且在按钮上抬起才发生该事件
'MouseMove--鼠标在按钮上移动即发生该事件
'Click--鼠标左键在按钮上按下并且在按钮上抬起才发生该事件
'KeyDown--键盘按键按下
'KeyPress--键盘普通键被敲击
'KeyUP--键盘按键抬起
'
'******************************谢谢您阅读本文件***********************************************

