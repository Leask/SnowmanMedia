VERSION 5.00
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "SmM_Tools.ocx"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѡ��"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5580
   FillColor       =   &H00FFFFFF&
   Icon            =   "����.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5580
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   5820
      Index           =   1
      Left            =   6030
      TabIndex        =   47
      Top             =   540
      Width           =   5550
      Begin VB.Frame Frame13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " ϵͳ "
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   135
         TabIndex        =   71
         Top             =   2025
         Width           =   5325
         Begin VB.CheckBox Check20 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�����̨���� Snowman Media ý������(&Y)"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   72
            Top             =   225
            Width           =   4380
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " �߼� "
         ForeColor       =   &H00000000&
         Height          =   1455
         Index           =   1
         Left            =   2115
         TabIndex        =   49
         Top             =   450
         Width           =   3345
         Begin VB.CheckBox Check7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�����ļ����Ҽ��˵�(&I)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   180
            TabIndex        =   55
            Top             =   495
            Value           =   1  'Checked
            Width           =   2985
         End
         Begin VB.CheckBox Check4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���� cda��dat �� vob �ļ�(&V)"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   53
            Top             =   1035
            Width           =   2985
         End
         Begin VB.CheckBox Check6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���� CD��VCD �� DVD ����(&D)"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   52
            Top             =   765
            Width           =   3030
         End
         Begin VB.CheckBox Check7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�����Զ��ָ�ý�����(&S)"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   51
            Top             =   225
            Width           =   2985
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   2250
            TabIndex        =   50
            Top             =   765
            Width           =   960
         End
      End
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " ��ʽ "
         ForeColor       =   &H00000000&
         Height          =   2895
         Index           =   1
         Left            =   135
         TabIndex        =   48
         Top             =   2790
         Width           =   5325
         Begin VB.ListBox List2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   2430
            IntegralHeight  =   0   'False
            ItemData        =   "����.frx":2CFA
            Left            =   180
            List            =   "����.frx":2D2B
            Style           =   1  'Checkbox
            TabIndex        =   56
            Top             =   270
            Width           =   4965
         End
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ʹ�ñ�ѡ����� Snowman Media ý�������"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   54
         Top             =   90
         Width           =   4110
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   2325
         Index           =   1
         Left            =   135
         Picture         =   "����.frx":3106
         Top             =   -135
         Width           =   1965
      End
   End
   Begin API���ƴ�ȫ.LyfTools LF1 
      Left            =   2160
      Top             =   7605
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   5820
      Left            =   0
      TabIndex        =   4
      Top             =   540
      Width           =   5550
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " ��ʾ "
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   135
         TabIndex        =   68
         Top             =   2790
         Width           =   5325
         Begin VB.CheckBox Check10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��ϵͳ������ʾ(&S)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   74
            Top             =   225
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox Check9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��С��ʱ����(&S)"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2520
            TabIndex        =   73
            Top             =   225
            Width           =   1995
         End
      End
      Begin VB.Frame Frame12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " ��� "
         ForeColor       =   &H00000000&
         Height          =   2130
         Left            =   135
         TabIndex        =   41
         Top             =   3555
         Width           =   5325
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   1650
            IntegralHeight  =   0   'False
            ItemData        =   "����.frx":52AB
            Left            =   180
            List            =   "����.frx":52FA
            Style           =   1  'Checkbox
            TabIndex        =   42
            Top             =   270
            Width           =   4965
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " ���� "
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   135
         TabIndex        =   11
         Top             =   2025
         Width           =   5325
         Begin VB.CheckBox Check19 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��ʾ��ӭ����(&W)"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   75
            Top             =   225
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " ���Ż� "
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   2115
         TabIndex        =   5
         Top             =   450
         Width           =   3345
         Begin VB.CheckBox Check5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��������Ψһʶ����Ĳ��Ż�(&N)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   9
            Top             =   1035
            Width           =   3030
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��������ʱ������������(&M)"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   8
            Top             =   225
            Width           =   2985
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�������Զ�����δ�������Ŀ(&L)"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   7
            Top             =   495
            Width           =   3030
         End
         Begin VB.CheckBox Check3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�����Ż��Զ�����(&U)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   6
            Top             =   765
            Width           =   2490
         End
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ʹ�ñ�ѡ����� Snowman Media �ĳ���ѡ�"
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   225
         TabIndex        =   10
         Top             =   90
         Width           =   4110
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   2325
         Left            =   135
         Picture         =   "����.frx":557A
         Top             =   -135
         Width           =   1965
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   5820
      Index           =   0
      Left            =   6300
      TabIndex        =   24
      Top             =   540
      Width           =   5550
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " �㲥 "
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   870
         Index           =   3
         Left            =   135
         TabIndex        =   64
         Top             =   4815
         Width           =   5325
         Begin VB.CheckBox Check6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�Զ�ѡ�����ÿ����������ڹ㲥(&P)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   3
            Left            =   180
            TabIndex        =   65
            Top             =   270
            Width           =   3660
         End
      End
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " ���� "
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1545
         Index           =   2
         Left            =   135
         TabIndex        =   57
         Top             =   3150
         Width           =   5325
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2430
            TabIndex        =   62
            Top             =   1080
            Width           =   870
         End
         Begin VB.CheckBox Check6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���ô˹����Ż�����(&T)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   180
            TabIndex        =   61
            Top             =   495
            Value           =   1  'Checked
            Width           =   3030
         End
         Begin VB.OptionButton Option15 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�ò��Ż��������������С(&E)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   450
            TabIndex        =   58
            Top             =   765
            Value           =   -1  'True
            Width           =   4515
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�Զ������Ʊ���(&B):"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   450
            TabIndex        =   59
            Top             =   1080
            Width           =   2040
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�ֽ�ÿ��"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   3420
            TabIndex        =   63
            Top             =   1125
            Width           =   1050
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����һ���������ƴ������ڱ�֤��������"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   225
            TabIndex        =   60
            Top             =   270
            Width           =   4290
         End
      End
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " ���� "
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1005
         Index           =   0
         Left            =   135
         TabIndex        =   28
         Top             =   2025
         Width           =   5325
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2520
            TabIndex        =   45
            Top             =   540
            Width           =   600
         End
         Begin VB.OptionButton Option15 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�ò��Ż������״�����ܾ�������ʱ��(&L)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   44
            Top             =   225
            Width           =   4155
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�Զ��建������СΪ(S):"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   29
            Top             =   540
            Width           =   2445
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   3240
            TabIndex        =   46
            Top             =   585
            Width           =   510
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " ���� "
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1455
         Index           =   0
         Left            =   2115
         TabIndex        =   25
         Top             =   450
         Width           =   3345
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   2250
            TabIndex        =   43
            Text            =   "163"
            Top             =   765
            Width           =   915
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����ָ������(U):"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   450
            TabIndex        =   67
            Top             =   765
            Width           =   2445
         End
         Begin VB.OptionButton Option15 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����Ĭ������(&D)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   450
            TabIndex        =   66
            Top             =   495
            Value           =   -1  'True
            Width           =   1680
         End
         Begin VB.CheckBox Check8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ʹ���Զ��Ͽ�(&O)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   450
            TabIndex        =   27
            Top             =   1035
            Width           =   2130
         End
         Begin VB.CheckBox Check7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�����Զ�����(&A)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   26
            Top             =   225
            Width           =   2985
         End
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ʹ�ñ�ѡ����� Snowma Media �����繦�ܡ�"
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   30
         Top             =   90
         Width           =   4110
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   2325
         Index           =   0
         Left            =   135
         Picture         =   "����.frx":73A8
         Top             =   -135
         Width           =   1965
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   5820
      Left            =   6255
      TabIndex        =   12
      Top             =   540
      Width           =   5550
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " DVD ���� "
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   1815
         Left            =   135
         TabIndex        =   31
         Top             =   3870
         Width           =   5325
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   300
            ItemData        =   "����.frx":9034
            Left            =   765
            List            =   "����.frx":9036
            TabIndex        =   40
            Text            =   "ϵͳĬ������"
            Top             =   1395
            Width           =   2760
         End
         Begin VB.OptionButton Option14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�ޱ�ʶ"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   450
            TabIndex        =   38
            Top             =   765
            Value           =   -1  'True
            Width           =   1050
         End
         Begin VB.OptionButton Option13 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "NC-17"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4050
            TabIndex        =   37
            Top             =   495
            Width           =   960
         End
         Begin VB.OptionButton Option12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "R"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3285
            TabIndex        =   36
            Top             =   495
            Width           =   825
         End
         Begin VB.OptionButton Option11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "PG-13"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2250
            TabIndex        =   35
            Top             =   495
            Width           =   915
         End
         Begin VB.OptionButton Option6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "PG"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1305
            TabIndex        =   34
            Top             =   495
            Width           =   915
         End
         Begin VB.CheckBox Check17 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�ҳ�����(&H)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   32
            Top             =   225
            Width           =   2985
         End
         Begin VB.OptionButton Option5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "P"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   450
            TabIndex        =   33
            Top             =   495
            Width           =   915
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���������(&P):"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   39
            Top             =   1125
            Width           =   1860
         End
      End
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " ��Ƶ "
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   2115
         TabIndex        =   18
         Top             =   450
         Width           =   3345
         Begin VB.CheckBox Check16 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�Զ���ֹ��Ļ����(&S)"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   22
            Top             =   765
            Width           =   2490
         End
         Begin VB.CheckBox Check15 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��󻯴���ʱ�Զ�����ȫ��ģʽ(&M)"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   21
            Top             =   495
            Width           =   3120
         End
         Begin VB.CheckBox Check14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������������ͣ��Ƶ(&C)"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   20
            Top             =   225
            Width           =   2985
         End
         Begin VB.CheckBox Check13 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������Ƶ����ʱ��Ч(&O)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   495
            TabIndex        =   19
            Top             =   1035
            Width           =   2580
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   " �б� "
         ForeColor       =   &H00000000&
         Height          =   1725
         Left            =   135
         TabIndex        =   13
         Top             =   2025
         Width           =   5325
         Begin VB.CheckBox Check18 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ָ���ļ�ʱ�������ļ���(&N)"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   69
            Top             =   1305
            Width           =   3795
         End
         Begin VB.CheckBox Check12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�˳�ʱ����б�(&A)"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   180
            TabIndex        =   17
            Top             =   225
            Width           =   2985
         End
         Begin VB.CheckBox Check11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�����Զ�����(L)"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   450
            TabIndex        =   16
            Top             =   495
            Width           =   3030
         End
         Begin VB.OptionButton Option8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����ʱ������������(&I)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   450
            TabIndex        =   15
            Top             =   765
            Width           =   2355
         End
         Begin VB.OptionButton Option7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�˳�ʱ������������(&X)"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   450
            TabIndex        =   14
            Top             =   1035
            Value           =   -1  'True
            Width           =   2310
         End
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ʹ�ñ�ѡ����� Snowma Media �������ص㡣"
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   225
         TabIndex        =   23
         Top             =   90
         Width           =   4110
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         Height          =   2325
         Left            =   135
         Picture         =   "����.frx":9038
         Top             =   -135
         Width           =   1965
      End
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4590
      MouseIcon       =   "����.frx":B141
      MousePointer    =   99  'Custom
      TabIndex        =   70
      Top             =   180
      Width           =   870
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1215
      MouseIcon       =   "����.frx":B293
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   180
      Width           =   1050
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ý��"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3465
      MouseIcon       =   "����.frx":B3E5
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   180
      Width           =   1050
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   330
      Left            =   4590
      Top             =   90
      Width           =   870
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2340
      MouseIcon       =   "����.frx":B537
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   180
      Width           =   1050
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Left            =   1215
      Top             =   90
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   90
      MouseIcon       =   "����.frx":B689
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   180
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Left            =   3465
      Top             =   90
      Width           =   1050
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Left            =   90
      Top             =   90
      Width           =   1050
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   330
      Left            =   2340
      Top             =   90
      Width           =   1050
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub SetList()
List1.Selected(0) = 1
List1.Selected(1) = 1
List1.Selected(2) = 1
List1.Selected(3) = 1
List1.Selected(4) = 1
List1.Selected(5) = 1
List1.Selected(6) = 1
List1.Selected(7) = 1
List1.Selected(8) = 1
List1.Selected(9) = 1
List1.Selected(10) = 1
List1.Selected(11) = 1
List1.Selected(12) = 1
List1.Selected(13) = 1
List1.Selected(14) = 1
List1.Selected(15) = 1
List1.Selected(16) = 0
List1.Selected(17) = 0
List1.Selected(18) = 1
List1.Selected(19) = 1
List1.Selected(20) = 0
List1.Selected(21) = 0
List1.Selected(22) = 0
List1.Selected(23) = 0
List1.Selected(24) = 0
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoStart", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoStart", False
End If
End Sub


Private Sub Check11_Click()
If Check11.Value = 1 Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoCln", True
Option7.Enabled = True
Option8.Enabled = True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoCln", False
Option7.Enabled = False
Option8.Enabled = False
End If
End Sub

Private Sub Check12_Click()
If Check12.Value = 1 Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Clean", True
Check11.Enabled = False
Check11.Value = 0
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoCln", False
Option7.Enabled = False
Option8.Enabled = False
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Clean", False
Check11.Enabled = True
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoCln") = True Then
Option7.Enabled = True
Option8.Enabled = True
End If
End If
End Sub

Private Sub Check13_Click()
If Check13.Value = 1 Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "OnlyVideo", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "OnlyVideo", False
End If
End Sub

Private Sub Check14_Click()
If Check14.Value = 1 Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "MousePu", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "MousePu", False
End If
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Change", True
End Sub

Private Sub Check15_Click()
If Check15.Value = 1 Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoFu", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoFu", False
End If
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Change", True
End Sub

Private Sub Check16_Click()
If Check16.Value = 1 Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ScSave", True
 Check13.Enabled = True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ScSave", False
 Check13.Enabled = False
End If
End Sub

Private Sub Check18_Click()
If Check18.Value = 1 Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AllFiles", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AllFiles", False
End If
End Sub

Private Sub Check19_Click()
If Check2.Value = 1 Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ShowPic", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ShowPic", False
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoCnt", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoCnt", False
End If
End Sub

Private Sub Check20_Click()
If Check20.Value = 1 Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "StartUp", True
LF1.PutToStartMenu App.Path + "\SmM_Helper.exe                                        "
Shell (App.Path + "\SmM_Helper.exe")
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "StartUp", False

End If

End Sub

Private Sub Check4_Click(Index As Integer)
If Index = 1 Then
If Check4(1).Value = 1 Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "DATVOB", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "DATVOB", False
End If
Shell (App.Path + "\SmM_Types.exe")
End If
End Sub


Private Sub Check6_Click(Index As Integer)
If Index = 1 Then
If Check6(1).Value = 1 Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "CVD", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "CVD", False
End If
Shell (App.Path + "\SmM_Types.exe")
End If
End Sub

Private Sub Check7_Click(Index As Integer)
If Index = 1 Then
If Check7(1).Value = 1 Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoMedia", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoMedia", False
End If
End If
Shell (App.Path + "\SmM_Types.exe")
'If Index = 2 Then
'If Check7(2).Value = 1 Then
'LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoDir", True
'Else
'LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoDir", False
'End If
'End If



End Sub

Private Sub Check9_Click()
If Check9.Value = 1 Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ShowRw", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ShowRw", False
End If
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Change", True
End Sub

Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance Then End
SetList
List1.ListIndex = 0
LF1.Addhorizon List2, 1360
List2.Selected(1) = True
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoStart") = True Then Check1.Value = 1
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoCnt") = True Then Check2.Value = 1
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "MousePu") = True Then Check14.Value = 1
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoFu") = True Then Check15.Value = 1
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ScSave") = True Then Check16.Value = 1
If Check16.Value = 1 Then Check13.Enabled = True
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "OnlyVideo") = True Then Check13.Value = 1
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Clean") = True Then Check12.Value = 1
If Check12.Value = 1 Then
Check11.Enabled = False
Check11.Value = 0
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoCln", False
Option7.Enabled = False
Option8.Enabled = False
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "Clean", False
Check11.Enabled = True
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoCln") = True Then
Option7.Enabled = True
Option8.Enabled = True
End If
End If
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoCln") = True Then Check11.Value = 1
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "StrCln") = True Then Option8.Value = True
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoMedia") = True Then Check7(1).Value = 1
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "CVD") = True Then Check6(1).Value = 1
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "DATVOB") = True Then Check4(1).Value = 1
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mwm") = True Then List2.Selected(2) = True
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mw") = True Then List2.Selected(3) = True
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Ivf") = True Then List2.Selected(4) = True
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Swf") = True Then List2.Selected(7) = True
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Aiff") = True Then List2.Selected(8) = True
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Rm") = True Then List2.Selected(5) = True
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Qt") = True Then List2.Selected(6) = True

If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mpeg") = True Then List2.Selected(9) = True
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Au") = True Then List2.Selected(10) = True
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mp3") = True Then List2.Selected(11) = True
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Midi") = True Then List2.Selected(12) = True
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ShowRw") = True Then Check9.Value = 1
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AllFiles") = True Then Check18.Value = 1
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "StartUp") = True Then Check20.Value = 1
'If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "AutoDir") = True Then Check7(2).Value = 1

If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "ShowPic") = True Then Check19.Value = 1


List1.ListIndex = 0

List2.ListIndex = 0

If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "RunTime") = 0 Then
Label3_MouseUp 1, 0, 0, 0
List2.ListIndex = 0
List2_Click
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "RunTime") = 0 Then
Shell App.Path + "\SmM_Types.exe"
If MsgBox("Snowman Media ilxz �Ѿ�׼��������Ҫ����������", vbYesNo) = vbYes Then Shell App.Path + "\Snowman.exe", vbNormalFocus
End If

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Frame3.Left = 0
Frame3.Visible = True
Frame3.Enabled = True
Frame4(1).Left = 10000
Frame4(1).Visible = False
Frame4(1).Enabled = False
Frame6.Left = 10000
Frame6.Visible = False
Frame6.Enabled = False
Frame4(0).Left = 10000
Frame4(0).Visible = False
Shape8.BackColor = &HFFC0C0
Shape2.BackColor = &HFFFFFF
Shape4.BackColor = &HFFFFFF
Shape1.BackColor = &HFFFFFF

End Sub

Private Sub Label11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Shell (App.Path + "\SmM_Help.exe")

End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Frame3.Left = 10000
Frame3.Visible = False
Frame3.Enabled = False
Frame4(1).Left = 10000
Frame4(1).Visible = False
Frame4(1).Enabled = False
Frame6.Left = 10000
Frame6.Visible = False
Frame6.Enabled = False
Frame4(0).Left = 0
Frame4(0).Visible = True
Shell "Rundll32.exe Shell32.dll,Control_RunDLL inetcpl.cpl"
Shape8.BackColor = &HFFFFFF
Shape2.BackColor = &HFFFFFF
Shape4.BackColor = &HFFC0C0
Shape1.BackColor = &HFFFFFF

End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Frame3.Left = 10000
Frame3.Visible = False
Frame3.Enabled = False
Frame4(1).Left = 0
Frame4(1).Visible = True
Frame4(1).Enabled = True
Frame6.Left = 10000
Frame6.Visible = False
Frame6.Enabled = False
Frame4(0).Left = 10000
Frame4(0).Visible = False
Shape8.BackColor = &HFFFFFF
Shape2.BackColor = &HFFFFFF
Shape4.BackColor = &HFFFFFF
Shape1.BackColor = &HFFC0C0

End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 1 Then Exit Sub
Frame3.Left = 10000
Frame3.Visible = False
Frame3.Enabled = False
Frame4(1).Left = 10000
Frame4(1).Visible = False
Frame4(1).Enabled = False
Frame6.Left = 0
Frame6.Visible = True
Frame6.Enabled = True
Frame4(0).Left = 10000
Frame4(0).Visible = False
Shape8.BackColor = &HFFFFFF
Shape2.BackColor = &HFFC0C0
Shape4.BackColor = &HFFFFFF
Shape1.BackColor = &HFFFFFF

End Sub



Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetList

End Sub

Private Sub List2_Click()
If List2.ListIndex = 1 Then
List2.Selected(1) = True
End If
If List2.ListIndex = 2 Then
If List2.Selected(2) = True Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mwm", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mwm", False
End If
End If
If List2.ListIndex = 3 Then
If List2.Selected(3) = True Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mw", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mw", False
End If
End If
If List2.ListIndex = 4 Then
If List2.Selected(4) = True Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Ivf", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Ivf", False
End If
End If
If List2.ListIndex = 5 Then
If List2.Selected(5) = True Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Rm", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Rm", False
End If
End If


If List2.ListIndex = 6 Then
If List2.Selected(6) = True Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Qt", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Qt", False
End If
End If
If List2.ListIndex = 7 Then
If List2.Selected(7) = True Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Swf", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Swf", False
End If
End If
If List2.ListIndex = 8 Then
If List2.Selected(8) = True Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Aiff", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Aiff", False
End If
End If
If List2.ListIndex = 9 Then
If List2.Selected(9) = True Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mpeg", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mpeg", False
End If
End If
If List2.ListIndex = 10 Then
If List2.Selected(10) = True Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Au", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Au", False
End If
End If
If List2.ListIndex = 11 Then
If List2.Selected(11) = True Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mp3", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mp3", False
End If
End If
If List2.ListIndex = 12 Then
If List2.Selected(12) = True Then
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Midi", True
Else
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Midi", False
End If
End If
If List2.ListIndex = 13 Then
List2.Selected(13) = False
End If
If List2.ListIndex = 0 Then
If List2.Selected(0) = True Then
List2.Selected(0) = False
List2.Selected(2) = True
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mwm", True
List2.Selected(3) = True
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mw", True
List2.Selected(4) = True
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Ivf", True
List2.Selected(5) = True
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Rm", True
List2.Selected(6) = True
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Qt", True
List2.Selected(7) = True
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Swf", True
List2.Selected(8) = True
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Aiff", True
List2.Selected(9) = True
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mpeg", True
List2.Selected(10) = True
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Au", True
List2.Selected(11) = True
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mp3", True
List2.Selected(12) = True
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Midi", True
List2.Selected(14) = False
End If
End If
If List2.ListIndex = 14 Then
If List2.Selected(14) = True Then
List2.Selected(0) = False
List2.Selected(2) = False
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mwm", False
List2.Selected(3) = False
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mw", False
List2.Selected(4) = False
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Ivf", False
List2.Selected(5) = False
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Rm", False
List2.Selected(6) = False
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Qt", False
List2.Selected(7) = False
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Swf", False
List2.Selected(8) = False
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Aiff", False
List2.Selected(9) = False
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mpeg", False
List2.Selected(10) = False
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Au", False
List2.Selected(11) = False
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mp3", False
List2.Selected(12) = False
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Midi", False
List2.Selected(14) = False
End If
End If
Shell (App.Path + "\SmM_Types.exe")
End Sub




Private Sub Option7_Click()
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "StrCln", False
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "OvrCln", True

End Sub

Private Sub Option8_Click()
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "StrCln", True
LF1.SetDWORDValue "HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "OvrCln", False
End Sub


