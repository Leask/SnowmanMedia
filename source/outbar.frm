VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#4.0#0"; "LYFTOOLS.OCX"
Object = "{628CC7D5-A6CF-11D0-B997-00805F024BFD}#1.0#0"; "VERTMENU.OCX"
Begin VB.Form outlook 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Snowman "
   ClientHeight    =   6435
   ClientLeft      =   705
   ClientTop       =   1860
   ClientWidth     =   1500
   ForeColor       =   &H00000000&
   Icon            =   "outbar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   1500
   ShowInTaskbar   =   0   'False
   Begin VertMenu.VerticalMenu VerticalMenu1 
      Height          =   6450
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "功能菜单  - Snowman Media  2.0"
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   11377
      MenusMax        =   6
      MenuCaption1    =   "Play Media"
      MenuItemsMax1   =   3
      MenuItemIcon11  =   "outbar.frx":1582
      MenuItemCaption11=   "Play Media"
      MenuItemIcon12  =   "outbar.frx":189C
      MenuItemCaption12=   "Play Flash"
      MenuItemIcon13  =   "outbar.frx":1BB6
      MenuItemCaption13=   "Play VCD"
      MenuCaption2    =   "Net Show"
      MenuItemsMax2   =   4
      MenuItemIcon21  =   "outbar.frx":1ED0
      MenuItemCaption21=   "Media on Line"
      MenuItemIcon22  =   "outbar.frx":21EA
      MenuItemCaption22=   "Flash on Line"
      MenuItemIcon23  =   "outbar.frx":2504
      MenuItemCaption23=   "Ring up"
      MenuItemIcon24  =   "outbar.frx":281E
      MenuItemCaption24=   "Ring off"
      MenuCaption3    =   "See Picture"
      MenuItemIcon31  =   "outbar.frx":2B38
      MenuItemCaption31=   "See Picture"
      MenuCaption4    =   "Play List"
      MenuItemsMax4   =   2
      MenuItemIcon41  =   "outbar.frx":2E52
      MenuItemCaption41=   "Play List"
      MenuItemIcon42  =   "outbar.frx":316C
      MenuItemCaption42=   "Edit List"
      MenuCaption5    =   "Music Box"
      MenuItemIcon51  =   "outbar.frx":3486
      MenuItemCaption51=   "Save Music"
      MenuCaption6    =   "Setting & Help"
      MenuItemsMax6   =   9
      MenuItemIcon61  =   "outbar.frx":37A0
      MenuItemCaption61=   "Aways on Top"
      MenuItemIcon62  =   "outbar.frx":3ABA
      MenuItemCaption62=   "EQ Setting"
      MenuItemIcon63  =   "outbar.frx":3DD4
      MenuItemCaption63=   "Hide Taskbar"
      MenuItemIcon64  =   "outbar.frx":40EE
      MenuItemCaption64=   "Show Taskbar"
      MenuItemIcon65  =   "outbar.frx":4408
      MenuItemCaption65=   "Snow"
      MenuItemIcon66  =   "outbar.frx":4722
      MenuItemCaption66=   "Help"
      MenuItemIcon67  =   "outbar.frx":4A3C
      MenuItemCaption67=   "About SmM"
      MenuItemIcon68  =   "outbar.frx":4D56
      MenuItemCaption68=   "SmM Web Side"
      MenuItemIcon69  =   "outbar.frx":5070
      MenuItemCaption69=   "Mail to Us"
   End
   Begin 刘玉锋的VB超级工具集.LyfTools LyfTools1 
      Height          =   480
      Left            =   4770
      TabIndex        =   0
      Top             =   5985
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4635
      Top             =   5085
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Files Opening Window  - Snowman Media  2.0"
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   465
      Left            =   4230
      TabIndex        =   1
      Top             =   2070
      Width           =   825
   End
End
Attribute VB_Name = "outlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Const HKEY_CURRENT_USER = &H80000001
Const ERROR_SUCCESS = 0&
Private Declare Function RasHangUp Lib "RasApi32.DLL" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
Private Declare Function RasEnumConnections Lib "RasApi32.DLL" Alias "RasEnumConnectionsA" (lprasconn As Any, lpcb As Long, lpcConnections As Long) As Long
Const RAS95_MaxEntryName = 256
Const RAS95_MaxDeviceName = 128
Const RAS_MaxDeviceType = 16
Private Type RASCONN95
dwSize As Long
hRasConn As Long
szEntryName(RAS95_MaxEntryName) As Byte
szDeviceType(RAS_MaxDeviceType) As Byte
szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
Public Function GetConnect() As String
    Dim hKey As Long
    Dim SubKey As String
    hKey = HKEY_CURRENT_USER
    SubKey = "RemoteAccess"
    GetConnect = GetRegValue(hKey, SubKey, "Default")
End Function
Public Function GetRegValue(hKey As Long, lpszSubKey As String, szKey As String) As Variant
    On Error GoTo ErrorRoutineErr:
    Dim phkResult As Long
    Dim lResult As Long
    Dim szBuffer As String
    Dim lBuffSize As Long
    szBuffer = Space(255)
    lBuffSize = Len(szBuffer)
    RegOpenKeyEx hKey, lpszSubKey, 0, 1, phkResult
   lResult = RegQueryValueEx(phkResult, szKey, 0, 0, szBuffer, lBuffSize)
   RegCloseKey phkResult
   If lResult = ERROR_SUCCESS Then
        GetRegValue = Left(szBuffer, lBuffSize - 1)
    Else
        GetRegValue = ""
    End If
    Exit Function
ErrorRoutineErr:
    GetRegValue = ""
End Function
Private Sub VerticalMenu1_MenuItemClick(MenuNumber As Long, MenuItem As Long)
 On Error Resume Next
Select Case MenuNumber
        Case 1
            Select Case MenuItem
                Case 1
                       outlook.CommonDialog1.Filter = "媒体文件:vcd、mp3、wma、wav、wax、asf、rmi、asx、mov、m1v、mp2、mpg、mpeg、mpa、mpe、avi、mid、qt、m3u、aif、aifc、aiff、au、snd..." & _
          "|*.au;*.dat;*.and;*.aif;*.aifc;*.aiff;*.mpe;*.mpa;*.wax;*.rmi;*.asx;*.m1v;*.mp2;*.asf;*.mov;*.mp3;*.qt;*.mpeg;*.mpg;*.m3u;*.wma;*.wav;*.avi;*.mid|浏览所有文件:*.*|*.*"
          outlook.CommonDialog1.FilterIndex = 1
          outlook.CommonDialog1.filename = ""
          outlook.CommonDialog1.ShowOpen
          If Len(outlook.CommonDialog1.filename) > 0 Then
         Form2.MediaPlayer1.filename = outlook.CommonDialog1.filename
          Form2.Caption = outlook.CommonDialog1.filename & "  - Snowman Media  2.0"
           Form2.Show
          End If
                Case 2
                   outlook.CommonDialog1.Filter = "Flash文件(*.swf)" & _
        "|*.swf|浏览所有文件:*.*|*.*"
        outlook.CommonDialog1.FilterIndex = 1
          outlook.CommonDialog1.filename = ""
          outlook.CommonDialog1.ShowOpen
          If Len(outlook.CommonDialog1.filename) > 0 Then
        Form7.Show
        Form7.ShockwaveFlash1.Movie = outlook.CommonDialog1.filename
           End If
        Case 3
        Form2.Caption = "Playing VCD  - Snowman Media  2.0"
         Form2.MediaPlayer1.filename = App.Path + "\smmvcd.m3u"
        Form2.Show
               End Select
        Case 2
            Select Case MenuItem
                Case 1
                      Form5.Show
                Case 2
                     Form1.Show
               Case 3
                 Shell "rundll rnaui.dll,RnaDial " + GetConnect, vbNormalFocus
               Case 4
               Dim lngRetCode As Long
Dim lpcb As Long
Dim lpcConnections As Long
Dim intArraySize As Integer
Dim intLooper As Integer
ReDim lprasconn95(intArraySize) As RASCONN95
lprasconn95(0).dwSize = 412
lpcb = 256 * lprasconn95(0).dwSize
lngRetCode = RasEnumConnections(lprasconn95(0), lpcb, lpcConnections)
If lngRetCode = 0 Then
If lpcConnections > 0 Then
For intLooper = 0 To lpcConnections - 1
RasHangUp lprasconn95(intLooper).hRasConn
Next intLooper
Else
MsgBox "没有拨号网络连接！", vbInformation
End If
End If
              End Select
        Case 3
            Select Case MenuItem
                Case 1
                         outlook.CommonDialog1.Filter = "图片文件:*.bmp、*.jpg、*.did、*.wmf、*.ico、*.gif、*.rle、*.cur、*.emf、*.png" & _
    "|*.bmp;*.jpg;*.did;*.wmf;*.ico;*.gif;*.rle;*.cur;*.emf;*.png|浏览所有文件:*.*|*.*"
       outlook.CommonDialog1.FilterIndex = 1
    outlook.CommonDialog1.ShowOpen
   If Len(outlook.CommonDialog1.filename) > 0 Then
    Form8.Image1.Picture = LoadPicture(outlook.CommonDialog1.filename)
    Form8.Image1.Picture = LoadPicture(outlook.CommonDialog1.filename)
    Form8.Caption = outlook.CommonDialog1.filename & "  - Snowman Media  2.0"
    Form8.Show
   End If
   End Select
        Case 4
            Select Case MenuItem
                Case 1
                            outlook.CommonDialog1.Filter = "列表文件:m3u" & _
          "|*.m3u|浏览所有文件:*.*|*.*"
    
          outlook.CommonDialog1.FilterIndex = 1
          outlook.CommonDialog1.filename = ""
          outlook.CommonDialog1.ShowOpen
          If Len(outlook.CommonDialog1.filename) > 0 Then
          Form2.Show
          Form2.MediaPlayer1.filename = outlook.CommonDialog1.filename
          End If
                Case 2
                     plm.Show
                End Select
           Case 5
            Select Case MenuItem
                Case 1
                If Len(Form102.TrackTime.Caption) > 0 Then
                   Form101.Show
                   Else
                   MsgBox ("光驱中没有CD,请先把CD放进光驱再使用本功能.")
                   End If
             End Select
             Case 6
            Select Case MenuItem
                Case 1
                      outlook.Label1.Caption = "0"
                       Form2.Timer1.Enabled = True
                Case 2
                 frmMain.Show
                 
                 Case 3
                   LyfTools1.HideTaskBar True
                Case 4
                  LyfTools1.HideTaskBar False
                Case 5
                     form100.Show
                    Case 6
                   Form3.Show
                   
                Case 7
                   Form4.Show
                   Case 8
                    LyfTools1.HttpTo "http://www.h2ont.com"
                    Case 9
                     LyfTools1.SendMail "h2ont@china.com"
            End Select
    End Select
End Sub
