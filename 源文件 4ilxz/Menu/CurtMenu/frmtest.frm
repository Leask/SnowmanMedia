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
   StartUpPosition =   1  '����������
   Begin CurtMenuǶ��ʽͼ�β˵�.CurtMenu CurtMenu1 
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
            Object.Tag             =   "�½�"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":0466
            Key             =   ""
            Object.Tag             =   "ѡ��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":05C2
            Key             =   ""
            Object.Tag             =   "��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":071E
            Key             =   ""
            Object.Tag             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":087A
            Key             =   ""
            Object.Tag             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":09D6
            Key             =   ""
            Object.Tag             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":0B32
            Key             =   ""
            Object.Tag             =   "ճ��"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmtest.frx":0C8E
            Key             =   ""
            Object.Tag             =   "�˳�"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuNew 
         Caption         =   "�½�"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "��"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "����"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveas 
         Caption         =   "���Ϊ(&A)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "�˳�"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuCut 
         Caption         =   "����"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "����"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "ճ��"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuOptions 
         Caption         =   "ѡ��(&O)"
         Begin VB.Menu mnuSubA 
            Caption         =   "�Ӳ˵�"
            Index           =   0
         End
         Begin VB.Menu mnuSubA 
            Caption         =   "�Ӳ˵�"
            Checked         =   -1  'True
            Index           =   1
         End
      End
      Begin VB.Menu mnuSub 
         Caption         =   "Ч��A"
         Index           =   0
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuSub 
         Caption         =   "Ч��B"
         Index           =   1
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuSub 
         Caption         =   "Ч��C"
         Index           =   2
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuSub 
         Caption         =   "Ч��D"
         Checked         =   -1  'True
         Index           =   3
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuSub 
         Caption         =   "Ч��E"
         Index           =   4
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuWWW 
         Caption         =   "�鿴���°汾"
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

'�Բ˵�����ת��
Private Sub Form_Load()
    CurtMenu1.Connect Me.hWnd, True, ImageList1
End Sub

'�����˵�����ʾ
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuEdit
    End If
End Sub

'�˵�Enabled���Ե���ʾ
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

'�˵�Checked���Ե���ʾ
Private Sub mnuSubA_Click(Index As Integer)
    mnuSubA(Index).Checked = Not mnuSubA(Index).Checked
End Sub


'�˵�Click�¼�����ʾ
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
'������վ��ȡ��������CurSoft��Ʒ
Private Sub mnuWWW_Click()
    ShellExecute 0&, "Open", "http://www.curtsoft.com", "", App.Path, 1
End Sub

'********************************�û���֪************************************************************
'
'   CurtMenu v1.01  ���������ʹ�ã�
'   ��Ȩ������:CurtSoft ����һ��Ȩ��!     http://www.curtsoft.com
'
'   ����ֻ���ʱ�����ڸ��������Ʒ�Ŀ�������ʹ�����Ծ�����������û�Э�飺
'   �����δ�����ؼ�������ҵĿ�ģ��������ʹ�ñ��ؼ��������븶�ѣ������û���39����λ�û���129��
'
'   ע�᷽������ע��Ѻ�ע����Ϣ�����ң��յ�ע��Э���ע��ɹ���
'   ��ϵ��ʽ��  Email��Inthenet@163.net      Mobile��13670102745     QQ��121728839
'   �����У���������������·���� �ʺţ�0755-36387681
'   ��ַ�������и�������·78�ŵ������Ĵ��ö�418  �ʱࣺ518031
'   ��ӭ���Ա��ؼ��������������ҽ����������ллʹ�ã�
'   ע�⣺������ע��Ѻ����õ����ʼ�����������ע����Ϣ���û�����֤�����롢��ϵ��ַ����
'
'*******************************���ܸ�Ҫ***************************************************************
'
'���ؼ��ص����£�
'    1-ʹ�ü�Ϊ�򵥣���һ��������Ч��Ƕ��VB�Դ��Ĳ˵��������ı���ʹ�÷�����
'    2-��ۿ��Ƽ����������Ч�����ޣ�
'    3-����WIN9X��WINNT��WIN2K��WINXP��
'
'******************************���¸��¼�¼*****************************************************************
'2002-02-26������Ϊv1.0.1����չ�������Ч�����Ľ��㷨���������ٶȡ�
'
'2002-02-25���ؼ��������汾v1.0.0
'
'******************************ʹ��˵��****************************************************************
'=================����==================
'
'�ı�/�ָ������в˵�����ʾЧ��:
'Connect(hWnd As Long, Flag As Boolean, Optional imlMenu As Object)
'    hWnd������Ҫ����ת���Ĳ˵��Ĵ���ľ����
'    Flag��TrueΪ����ת����FalseΪ���ת����
'    imlMenu�������˵�Ҫʹ��ͼ���ImageList�ؼ������ת����ʡ�Ըò�������
'ע�ͣ�Ҫ���˵�ITEM��ICON���й����������ò˵���ı��⣨Caption����ImageList�ؼ���ICON�ı�ǣ�Tag����ͬ��
'==================����=================
'BackColor����/���ò˵����ı�����ɫ
'ForeColor����/���ò˵����ֵ���ɫ
'IconBarColor����/����ͼ�����ı�����ɫ
'TextBarColor����/�����������ı�����ɫ
'ShadowColor����/������Ӱ��ɫ
'DisabledColor����/���ò˵�����õ���ɫ
'CheckMarkColor����/���ò˵����ѡʱ����ɫ
'HoverFillStyle����/���ò˵��������ʾ���������
'HoverEdgeColor����/���ò˵��������ʾ�ı߿���ɫ
'HoverForeColor�˵��������ʾ��������ɫ
'SepraterColor����/���÷ָ��ߵ���ɫ
'
'******************************лл���Ķ����ļ�***********************************************
