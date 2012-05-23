VERSION 5.00
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Object = "{244E6785-6684-11D2-943F-A976CFB4FC0C}#1.0#0"; "CTLSTBAR.OCX"
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "LYFTOOLS.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sm.M. Groupware"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   Icon            =   "Other.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   7035
   StartUpPosition =   2  '屏幕中心
   Begin CTLISTBARLibCtl.ctListBar ctListBar1 
      Height          =   1365
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "功能菜单"
      Top             =   0
      Width           =   6990
      _Version        =   65536
      _ExtentX        =   12330
      _ExtentY        =   2408
      _StockProps     =   70
      Caption         =   "已经安装"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackImage       =   "Other.frx":1582
      ListForeColor   =   16711680
      BarForeColor    =   16711680
      ListBarStyle    =   0
      BarHeight       =   21
      WordWrap        =   -1  'True
      Caption         =   "已经安装"
      PicArray0       =   "Other.frx":4647
      PicArray1       =   "Other.frx":4961
      PicArray2       =   "Other.frx":497D
      PicArray3       =   "Other.frx":4999
      PicArray4       =   "Other.frx":49B5
      PicArray5       =   "Other.frx":49D1
      PicArray6       =   "Other.frx":49ED
      PicArray7       =   "Other.frx":4A09
      PicArray8       =   "Other.frx":4A25
      PicArray9       =   "Other.frx":4A41
      PicArray10      =   "Other.frx":4A5D
      PicArray11      =   "Other.frx":4A79
      PicArray12      =   "Other.frx":4A95
      PicArray13      =   "Other.frx":4AB1
      PicArray14      =   "Other.frx":4ACD
      PicArray15      =   "Other.frx":4AE9
      PicArray16      =   "Other.frx":4B05
      PicArray17      =   "Other.frx":4B21
      PicArray18      =   "Other.frx":4B3D
      PicArray19      =   "Other.frx":4B59
      PicArray20      =   "Other.frx":4B75
      PicArray21      =   "Other.frx":4B91
      PicArray22      =   "Other.frx":4BAD
      PicArray23      =   "Other.frx":4BC9
      PicArray24      =   "Other.frx":4BE5
      PicArray25      =   "Other.frx":4C01
      PicArray26      =   "Other.frx":4C1D
      PicArray27      =   "Other.frx":4C39
      PicArray28      =   "Other.frx":4C55
      PicArray29      =   "Other.frx":4C71
      PicArray30      =   "Other.frx":4C8D
      PicArray31      =   "Other.frx":4CA9
      PicArray32      =   "Other.frx":4CC5
      PicArray33      =   "Other.frx":4CE1
      PicArray34      =   "Other.frx":4CFD
      PicArray35      =   "Other.frx":4D19
      PicArray36      =   "Other.frx":4D35
      PicArray37      =   "Other.frx":4D51
      PicArray38      =   "Other.frx":4D6D
      PicArray39      =   "Other.frx":4D89
      PicArray40      =   "Other.frx":4DA5
      PicArray41      =   "Other.frx":4DC1
      PicArray42      =   "Other.frx":4DDD
      PicArray43      =   "Other.frx":4DF9
      PicArray44      =   "Other.frx":4E15
      PicArray45      =   "Other.frx":4E31
      PicArray46      =   "Other.frx":4E4D
      PicArray47      =   "Other.frx":4E69
      PicArray48      =   "Other.frx":4E85
      PicArray49      =   "Other.frx":4EA1
      PicArray50      =   "Other.frx":4EBD
      PicArray51      =   "Other.frx":4ED9
      PicArray52      =   "Other.frx":4EF5
      PicArray53      =   "Other.frx":4F11
      PicArray54      =   "Other.frx":4F2D
      PicArray55      =   "Other.frx":4F49
      PicArray56      =   "Other.frx":4F65
      PicArray57      =   "Other.frx":4F81
      PicArray58      =   "Other.frx":4F9D
      PicArray59      =   "Other.frx":4FB9
      PicArray60      =   "Other.frx":4FD5
      PicArray61      =   "Other.frx":4FF1
      PicArray62      =   "Other.frx":500D
      PicArray63      =   "Other.frx":5029
      PicArray64      =   "Other.frx":5045
      PicArray65      =   "Other.frx":5061
      PicArray66      =   "Other.frx":507D
      PicArray67      =   "Other.frx":5099
      PicArray68      =   "Other.frx":50B5
      PicArray69      =   "Other.frx":50D1
      PicArray70      =   "Other.frx":50ED
      PicArray71      =   "Other.frx":5109
      PicArray72      =   "Other.frx":5125
      PicArray73      =   "Other.frx":5141
      PicArray74      =   "Other.frx":515D
      PicArray75      =   "Other.frx":5179
      PicArray76      =   "Other.frx":5195
      PicArray77      =   "Other.frx":51B1
      PicArray78      =   "Other.frx":51CD
      PicArray79      =   "Other.frx":51E9
      PicArray80      =   "Other.frx":5205
      PicArray81      =   "Other.frx":5221
      PicArray82      =   "Other.frx":523D
      PicArray83      =   "Other.frx":5259
      PicArray84      =   "Other.frx":5275
      PicArray85      =   "Other.frx":5291
      PicArray86      =   "Other.frx":52AD
      PicArray87      =   "Other.frx":52C9
      PicArray88      =   "Other.frx":52E5
      PicArray89      =   "Other.frx":5301
      PicArray90      =   "Other.frx":531D
      PicArray91      =   "Other.frx":5339
      PicArray92      =   "Other.frx":5355
      PicArray93      =   "Other.frx":5371
      PicArray94      =   "Other.frx":538D
      PicArray95      =   "Other.frx":53A9
      PicArray96      =   "Other.frx":53C5
      PicArray97      =   "Other.frx":53E1
      PicArray98      =   "Other.frx":53FD
      PicArray99      =   "Other.frx":5419
   End
   Begin API控制大全.LyfTools LyfTools1 
      Left            =   1485
      Top             =   495
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin ACTIVESKINLibCtl.SkinForm SkinForm1 
      Height          =   480
      Left            =   2835
      OleObjectBlob   =   "Other.frx":5435
      TabIndex        =   1
      Top             =   405
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub ctListBar1_ItemClick(ByVal nList As Integer, ByVal nItem As Integer)
On Error Resume Next
If nList = 2 Then
If Me.LyfTools1.IsConnected = True Then
Me.LyfTools1.HttpTo ("http://www.h2ont.com/snowmanmedia/groupware.htm")
Else
MsgBox ("无法连接流动网络 Sm.M. 插件中心,请确认连接网络后重试.")
End If
Exit Sub
End If
Shell LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "GwGn_" + Str(nItem))
End Sub

Private Sub Form_Load()
On Error Resume Next
If Val(LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "GwCt")) = 0 Then
MsgBox ("找不到已经安装的外加插件,请下载功能插件并安装后重试.")
End
End If
SkinForm1.SkinPath = LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path")
ctListBar1.AddList "更多插件"
ctListBar1.AddListImage 2, "下载插件", 1
 Dim i As Integer
For i = 1 To Val(LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "GwCt"))
ctListBar1.AddListImage 1, LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "GwNm_" + Str(i)), i + 1
Next
End Sub
