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
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   240
      ScaleHeight     =   825
      ScaleWidth      =   3885
      TabIndex        =   8
      Top             =   840
      Width           =   3885
      Begin CurtButton����ť�ؼ�.CurtButton CurtButton2 
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
      Begin CurtButton����ť�ؼ�.CurtButton CurtButton2 
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
      Begin CurtButton����ť�ؼ�.CurtButton CurtButton2 
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
      Begin CurtButton����ť�ؼ�.CurtButton CurtButton2 
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
      Begin CurtButton����ť�ؼ�.CurtButton CurtButton1 
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
      Begin CurtButton����ť�ؼ�.CurtButton CurtButton1 
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
      Begin CurtButton����ť�ؼ�.CurtButton CurtButton1 
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
      Begin CurtButton����ť�ؼ�.CurtButton CurtButton1 
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
      Begin CurtButton����ť�ؼ�.CurtButton CurtButton1 
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
   Begin CurtButton����ť�ؼ�.CurtButton CurtButton4 
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
   Begin CurtButton����ť�ؼ�.CurtButton CurtButton3 
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
   Begin CurtButton����ť�ؼ�.CurtButton CurtButton5 
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

    'bt3D�͵İ�ť������OFFICE97���͵Ĺ�����
    For i = 0 To 4
        CurtButton1(i).Appearance = bt3D
        CurtButton1(i).Move 4 + i * 24, 2, 24, 24
        CurtButton1(i).ToolTipText = "bt3D"
        Set CurtButton1(i).Picture = LoadPicture(App.Path & "\" & CStr(i + 1) & ".ico")
    Next
    'btXP�͵İ�ť������XP�͵Ĺ�����
    For i = 0 To 3
        CurtButton2(i).Appearance = btXP
        CurtButton2(i).Move 4 + i * 58, 2, 58, 50
        CurtButton2(i).ToolTipText = "btXP"
        Set CurtButton2(i).Picture = LoadPicture(App.Path & "\" & CStr(i + 6) & ".ico")
    Next
    'btXPplus�͵İ�ť�������ڰ�ť
    With CurtButton3
        .Appearance = btXPplus
        .ShowFocus = True
        .Default = True
        .ToolTipText = "btXPplus"
        .HoverFillStyle = hsColumn
        .Caption = "ȷ��"
    End With
    With CurtButton4
        .Appearance = btXPplus
        .ShowFocus = True
        .Cancel = True
        .ToolTipText = "btXPplus"
        .HoverFillStyle = hsLtlByLtl
        .Caption = "ȡ��"
    End With
    'btLabel�͵İ�ť�������ڱ�ǩ
    With CurtButton5
        .Appearance = btLabel
        .Alignment = alnCenterMiddle
        .ToolTipText = "btLabel"
        .Caption = "http://www.curtsoft.com"
    End With
End Sub
'���ؼ������ṩ��MouseEnter��MouseLeave�¼��������󷽱������
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

'��ʾ��ݼ���DEFAULT��CANCEL����
Private Sub CurtButton3_Click()
    MsgBox "лл�����ñ��ؼ���"
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
'********************************�û���֪************************************************************
'
'   CurtButton v1.04  ���������ʹ�ã�
'   ��Ȩ������:CurtSoft ����һ��Ȩ��!     http://www.curtsoft.com
'
'   ����ֻ���ʱ�����ڸ��������Ʒ�Ŀ�������ʹ�����Ծ�����������û�Э�飺
'   �����δ�����ؼ�������ҵĿ�ģ��������ʹ�ñ��ؼ��������븶�ѣ������û���29����λ�û���99��
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
'    1-ӵ��XPplus���ʺ�����ť����XP��3D���ʺ�������������Label��������ǩ�����ַ��
'    2-�ṩ�������¼�����(MouseEnter,MouseLeave�ȣ������󷽱���(���¼�˵������
'    3-XP��XPplus������ӰЧ���������ã�ͨ���������ԾͿ������ư�ť��ۡ�
'
'******************************���¸��¼�¼*****************************************************************
'
'���ؼ��ص����£�
'    1-ӵ�ж��ַ�񣬿����ڰ�ť������������ǩ�ȣ���;�㷺��
'    2-�ṩ�������¼����̣�����MouseLeave�����󷽱��̣�
'    3-��ʾЧ���������ã���ۿ��Ƽ���
'
'******************************���¸��¼�¼*****************************************************************
'2002-02-25�����������Ч������ΪhsLtlByLtlʱӦ�ó����޷�������ʾ�����⣻
'
'2002-02-17:����HoverFillStyle���ԣ���ťЧ�����ӷḻ��
'           �����˵����ť�����Ի����ť�޷���ԭ��BUG��
'
'2002-01-28���ؼ��������汾v1.0.0
'
'******************************ʹ��˵��****************************************************************
'==================����=================
'
'���з��ť��ʹ�õ����ԣ�
'Appearance--����/���ð�ť���
'Caption--����/��ʾ��ť�ı�
'Picture--����/��ʾ��ť��ʾ��ͼƬ��btLabel��񽫺��Ը����ԣ�
'BackColor--����/���ð�ť�ı�����ɫ
'ForeColor--����/���ð�ť��ǰ����ɫ
'Font--����/������ʾ�ı�ʹ�õ�����
'Enabled--����/���ð�ť�Ƿ����
'MouseIcon--����/���ð�ť���Զ������
'MousePointer--����/���ð�ť��ϵͳ���
'Cancel--����/���ð�ť�Ƿ�Ϊ����ġ�ȡ������ť
'Default--����/���ð�ť�Ƿ�Ϊ�����ȱʡ��ť
'
'XP��XPplus���ťʹ�õ����ԣ�
'HoverFillStyle����/��������ڰ�ť��ʱ���������
'HoverColor--����/��ʾ����ƶ�����ť�ϵ������ɫ
'MouseDownColor--����/������������ʱ�������ɫ
'EdgeColor--����/��ʾ����ƶ�����ť�ϵı߿���ɫ
'ShadowOffSet--����/����ͼ�����Ӱ��λ����
'ShadowColor--����/������Ӱ��ɫ
'
'XPplus���ťʹ�õ����ԣ�
'BorderColor����/��ʾ����Ƴ�����ťʱ�ı߿���ɫ
'ShowFocus����/���ð�ť�Ƿ��ڻ�ý���ʱ��ʾ����
'
'btLabel���ťʹ�õ����ԣ�
'Alignment����/�����ı��Ķ��뷽ʽ
'
'=================�¼�==================
'MouseEnter--�����밴ťʱ����
'MouseLeave--����뿪��ťʱ����
'MouseDown--���������ڰ�ť�ϰ��¶��������¼��������δUP���뿪��ť�����¼�����������
'MouseUp--���������ڰ�ť�ϰ��²����ڰ�ť��̧��ŷ������¼�
'MouseMove--����ڰ�ť���ƶ����������¼�
'Click--�������ڰ�ť�ϰ��²����ڰ�ť��̧��ŷ������¼�
'KeyDown--���̰�������
'KeyPress--������ͨ�����û�
'KeyUP--���̰���̧��
'
'******************************лл���Ķ����ļ�***********************************************

