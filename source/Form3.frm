VERSION 5.00
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#4.0#0"; "LYFTOOLS.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form3 
   Caption         =   "Help Window  - Snowman Media  2.0"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   11880
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   8070
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   14235
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "What's SmM"
      TabPicture(0)   =   "Form3.frx":1582
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "What's New"
      TabPicture(1)   =   "Form3.frx":159E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image2"
      Tab(1).Control(1)=   "Image3"
      Tab(1).Control(2)=   "Text2"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "How to Use"
      TabPicture(2)   =   "Form3.frx":15BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Instruction"
      TabPicture(3)   =   "Form3.frx":15D6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Image5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Text3"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Help on Line"
      TabPicture(4)   =   "Form3.frx":15F2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Image6"
      Tab(4).Control(1)=   "Image7"
      Tab(4).Control(2)=   "Image8"
      Tab(4).Control(3)=   "Image9"
      Tab(4).Control(4)=   "Image10"
      Tab(4).Control(5)=   "Image11"
      Tab(4).Control(6)=   "Image12"
      Tab(4).Control(7)=   "Label3"
      Tab(4).Control(8)=   "Label4"
      Tab(4).Control(9)=   "LyfTools1"
      Tab(4).ControlCount=   10
      Begin 刘玉锋的VB超级工具集.LyfTools LyfTools1 
         Height          =   480
         Left            =   -25000
         TabIndex        =   8
         Top             =   990
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   5685
         Left            =   -69645
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "Form3.frx":160E
         Top             =   1485
         Width           =   5145
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   4425
         Left            =   -71040
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "Form3.frx":18CF
         Top             =   2700
         Width           =   4875
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   3525
         Left            =   3420
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "Form3.frx":1BEC
         Top             =   2970
         Width           =   4830
      End
      Begin VB.Label Label4 
         Caption         =   "在 Snowman Media  2.0 的使用上遇到任何疑问请联系流动网络H2ont..."
         Height          =   555
         Left            =   -67305
         TabIndex        =   7
         Top             =   7065
         Width           =   4110
      End
      Begin VB.Label Label3 
         Caption         =   "想获得关于 Snowman Media  2.0 更多帮助请进入流动网络H2ont在线帮助..."
         Height          =   585
         Left            =   -67305
         TabIndex        =   6
         Top             =   5940
         Width           =   3960
      End
      Begin VB.Image Image12 
         Height          =   480
         Left            =   -68070
         MouseIcon       =   "Form3.frx":1C53
         MousePointer    =   99  'Custom
         Picture         =   "Form3.frx":1DA5
         ToolTipText     =   "h2ont@china.com  - Snowman Mdeia  2.0"
         Top             =   6705
         Width           =   480
      End
      Begin VB.Image Image11 
         Height          =   480
         Left            =   -68070
         MouseIcon       =   "Form3.frx":20AF
         MousePointer    =   99  'Custom
         Picture         =   "Form3.frx":2201
         ToolTipText     =   "http://www.h2ont.com/smm/help.html  - Snowman Media  2.0"
         Top             =   5625
         Width           =   480
      End
      Begin VB.Image Image10 
         Height          =   1035
         Left            =   -67980
         Picture         =   "Form3.frx":250B
         Top             =   585
         Width           =   3480
      End
      Begin VB.Image Image9 
         Height          =   7335
         Left            =   -74730
         Picture         =   "Form3.frx":3AF9
         Top             =   810
         Width           =   4965
      End
      Begin VB.Image Image8 
         Height          =   1500
         Left            =   -66585
         Picture         =   "Form3.frx":98B5
         Top             =   3285
         Width           =   1500
      End
      Begin VB.Image Image7 
         Height          =   1500
         Left            =   -65325
         Picture         =   "Form3.frx":B9D4
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Image Image6 
         Height          =   1530
         Left            =   -67845
         Picture         =   "Form3.frx":D4CA
         Top             =   2115
         Width           =   1545
      End
      Begin VB.Image Image5 
         Height          =   6780
         Left            =   -71760
         Picture         =   "Form3.frx":E703
         Top             =   1035
         Width           =   1530
      End
      Begin VB.Image Image4 
         Height          =   7200
         Left            =   -71625
         Picture         =   "Form3.frx":10708
         Top             =   585
         Width           =   5910
      End
      Begin VB.Image Image3 
         Height          =   1500
         Left            =   -73110
         Picture         =   "Form3.frx":188A3
         Top             =   4545
         Width           =   1500
      End
      Begin VB.Image Image2 
         Height          =   1500
         Left            =   -73110
         Picture         =   "Form3.frx":1A555
         Top             =   2250
         Width           =   1500
      End
      Begin VB.Label Label2 
         Height          =   3255
         Left            =   1575
         TabIndex        =   2
         Top             =   2610
         Width           =   4605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Snowman Media 2.0是一个Windows9x/ME下的多媒体播放软件。"
         Height          =   180
         Left            =   3420
         TabIndex        =   1
         Top             =   2160
         Width           =   4950
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   1980
         Picture         =   "Form3.frx":1C19E
         Top             =   1575
         Width           =   825
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer
Private Sub Form_Load()
a = 1
End Sub
Private Sub form_resize()
If a = 1 Then
Form3.SSTab1.Height = Form3.Height - 400
Form3.SSTab1.Width = Form3.ScaleWidth - 20
End If
a = a + 1
End Sub
Private Sub Image11_Click()
LyfTools1.HttpTo "http://www.h2ont.com/smm/help.html"
End Sub
Private Sub Image12_Click()
 LyfTools1.SendMail "h2ont@china.com"
End Sub
