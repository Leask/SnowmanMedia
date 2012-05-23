VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Formjp 
   Caption         =   "Snowman Media Screen Copier  1.0"
   ClientHeight    =   5310
   ClientLeft      =   2055
   ClientTop       =   2970
   ClientWidth     =   7245
   Icon            =   "截屏.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5310
   ScaleWidth      =   7245
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   7350
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10320
      Begin VB.VScrollBar VScroll1 
         Height          =   2535
         Left            =   2880
         TabIndex        =   4
         Top             =   0
         Width           =   195
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   2520
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         FillStyle       =   4  'Upward Diagonal
         Height          =   2535
         Left            =   0
         ScaleHeight     =   2475
         ScaleWidth      =   2835
         TabIndex        =   1
         Top             =   0
         Width           =   2895
         Begin VB.PictureBox picCopy 
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   0
            ScaleHeight     =   855
            ScaleWidth      =   1335
            TabIndex        =   2
            Top             =   0
            Width           =   1335
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3600
         Top             =   960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   7035
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7260
      Begin VB.PictureBox picScroll 
         Height          =   4920
         Left            =   0
         ScaleHeight     =   4860
         ScaleWidth      =   5175
         TabIndex        =   10
         Top             =   630
         Width           =   5235
         Begin VB.PictureBox picBmp 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   2805
            Left            =   0
            ScaleHeight     =   2805
            ScaleWidth      =   3840
            TabIndex        =   11
            Top             =   0
            Width           =   3840
         End
      End
      Begin VB.Timer tmr 
         Enabled         =   0   'False
         Interval        =   400
         Left            =   6750
         Top             =   5535
      End
      Begin VB.PictureBox picSnap 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   0
         Picture         =   "截屏.frx":1582
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         Top             =   45
         Width           =   540
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   4920
         Left            =   5220
         TabIndex        =   8
         Top             =   630
         Width           =   195
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   5535
         Width           =   6090
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存图片(&S)"
         Height          =   375
         Left            =   2700
         TabIndex        =   6
         Top             =   6165
         Width           =   1230
      End
      Begin MSComDlg.CommonDialog cdlg 
         Left            =   6615
         Top             =   6165
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "左键点击 Snowman Media  3.0 的图标然后按住不放拖动鼠标到需要捕获的窗口再松开即可截取所需窗体内容."
         Height          =   675
         Left            =   630
         TabIndex        =   12
         Top             =   45
         Width           =   4230
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Menu memFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu Save 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "退出(&X)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mDx 
      Caption         =   "对象(&V)"
      Begin VB.Menu Full 
         Caption         =   "全屏幕抓取(&F)"
         Checked         =   -1  'True
         Shortcut        =   ^F
      End
      Begin VB.Menu UsWin 
         Caption         =   "使用中窗口(&U)"
         Shortcut        =   ^U
      End
      Begin VB.Menu DiWin 
         Caption         =   "自定义窗口(&D)"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu Tool 
      Caption         =   "选项(&T)"
      Begin VB.Menu MtWin 
         Caption         =   "自定义窗口抓图时隐藏本窗口(&H)"
         Checked         =   -1  'True
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "Formjp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function GetROP2 Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDCEx Lib "user32" (ByVal hWnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Dim SnapHwnd&
Dim DeskHwnd&, DeskDC&
Dim oldRop2&
Dim rc As RECT
Dim meStates As Long
Const uID = 9998
Const uMessage = WM_USERB + 100
Sub SetPicture()
   picCopy.Visible = True
   If picCopy.Width <= Picture1.ScaleWidth Then
       picCopy.Left = (Picture1.ScaleWidth - picCopy.Width) / 2
   Else
       picCopy.Left = 0
       HScroll1.Min = 0
       HScroll1.Value = 0
       HScroll1.Max = picCopy.Width - Picture1.ScaleWidth
       HScroll1.SmallChange = IIf(HScroll1.Max \ 100 > 0, HScroll1.Max \ 100, 1)
       HScroll1.LargeChange = IIf(HScroll1.Max \ 10 > 0, HScroll1.Max \ 10, 1)
   End If
   If picCopy.Height <= Picture1.ScaleHeight Then
       picCopy.Top = (Picture1.ScaleHeight - picCopy.Height) / 2
   Else
       picCopy.Top = 0
       VScroll1.Min = 0
       VScroll1.Value = 0
       VScroll1.Max = picCopy.Height - Picture1.ScaleHeight
       VScroll1.SmallChange = IIf(VScroll1.Max \ 100 > 0, VScroll1.Max \ 100, 1)
       VScroll1.LargeChange = IIf(VScroll1.Max \ 10 > 0, VScroll1.Max \ 10, 1)
   End If
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
 On Error GoTo ErrMsg
    SetKeyboardHook Me.hWnd, WM_USERB
    prevWndProc = GetWindowLong(Me.hWnd, GWL_WNDPROC)
    SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf WndProc
    Dim nID As NOTIFYICONDATA
    nID.cbSize = Len(nID)
    nID.hWnd = Me.hWnd
    nID.uID = uID
    nID.uFlags = NIF_ICON + NIF_TIP + NIF_MESSAGE
    nID.hIcon = Me.Icon
    nID.szTip = "Snowman Media Screen Copier  1.0" + Chr(0)
    nID.uCallbackMessage = uMessage
    Shell_NotifyIcon NIM_ADD, nID
    Exit Sub
ErrMsg:
    MsgBox "Keybhook.dll 文件发生错误,无法运行.请重新启动 SnowmanMedia Screen Copier  1.0."
    End
End Sub
Private Sub Form_Resize()
    Frame1.Width = 30000
    Frame1.Height = 30000
    Frame2.Width = 30000
    Frame2.Height = 30000
    On Error Resume Next
    If Me.WindowState = vbMinimized Then
        Me.Hide
        Exit Sub
    Else
        Picture1.Width = Me.ScaleWidth - VScroll1.Width
        Picture1.Height = Me.ScaleHeight - HScroll1.Height
        VScroll1.Left = Picture1.Width
        HScroll1.Top = Picture1.Height
        VScroll1.Height = Picture1.Height
        HScroll1.Width = Picture1.Width
        SetPicture
    End If
        On Local Error Resume Next
    If WindowState <> vbMinimized Then
        meStates = Me.WindowState
       picScroll.Width = Me.ScaleWidth - 200
        picScroll.Height = Me.ScaleHeight - 830
        HScroll2.Top = picScroll.Height + 640
        HScroll2.Width = picScroll.Width
        VScroll2.Left = picScroll.Left + picScroll.Width
        VScroll2.Height = picScroll.Height
        If picScroll.Width > picBmp.Width Then
            HScroll2.Visible = False
        Else
            HScroll2.Visible = True
            HScroll2.Value = 0
            HScroll2.Max = picBmp.Width - picScroll.Width + 60
            HScroll2.LargeChange = picScroll.Width \ 3
            HScroll2.SmallChange = Screen.TwipsPerPixelX
            If HScroll2.LargeChange = 0 Then HScroll2.LargeChange = HScroll2.SmallChange
        End If
        If picScroll.Height > picBmp.Height Then
            VScroll2.Visible = False
        Else
            VScroll2.Visible = True
            VScroll2.Value = 0
            VScroll2.Max = picBmp.Height - picScroll.Height + 60
            VScroll2.LargeChange = picScroll.Height \ 3
            VScroll2.SmallChange = Screen.TwipsPerPixelY
            If VScroll2.LargeChange = 0 Then VScroll2.LargeChange = VScroll2.SmallChange
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ReleaseKeyboardHook
    SetWindowLong Me.hWnd, GWL_WNDPROC, prevWndProc
    Dim nID As NOTIFYICONDATA
    nID.cbSize = Len(nID)
    nID.hWnd = Me.hWnd
    nID.uID = uID
    Shell_NotifyIcon NIM_DELETE, nID
    mfrmMain.Show
End Sub
Private Sub Full_Click()
Full.Checked = True
UsWin.Checked = False
DiWin.Checked = False
    Frame1.Visible = True
    Frame2.Visible = False
    Me.WindowState = vbMinimized
     MsgBox "准备好后按下 F7 即可抓取图像."
End Sub
Private Sub MtWin_Click()
If MtWin.Checked = False Then
MtWin.Checked = True
Exit Sub
End If
If MtWin.Checked = True Then
MtWin.Checked = False
Exit Sub
End If
End Sub

Private Sub Save_Click()
If Me.Full.Checked = True Or Me.UsWin.Checked = True Then
      On Error Resume Next
      With CommonDialog1
      .DialogTitle = "另存为"
      .Filter = "位图文件:Bmp|*.bmp"
      .CancelError = True
      .ShowOpen
      If Err.Number <> cdlCancel Then
         SavePicture picCopy.Image, .FileName
      End If
   End With
Else
      cdlg.InitDir = App.Path
    cdlg.Filter = "位图文件:Bmp|*.Bmp"
    cdlg.ShowSave
    If Len(cdlg.FileName) = 0 Then Exit Sub
    Dim msg$
    msg$ = vbYes
    If Dir(cdlg.FileName) <> "" Then
       msg$ = MsgBox("文件已经存在,是否覆盖文件?覆盖后被覆盖的文件将无法恢复.", vbYesNo + vbQuestion, "询问")
    End If
    Select Case msg
        Case vbYes
            VB.SavePicture picBmp.Image, cdlg.FileName
        Case vbNo
    End Select
End If
End Sub

Private Sub UsWin_Click()
UsWin.Checked = True
Full.Checked = False
DiWin.Checked = False
     Frame1.Visible = True
    Frame2.Visible = False
 Me.WindowState = vbMinimized
   MsgBox "准备好后按下 F7 即可抓取图像."
End Sub
Private Sub DiWin_Click()
DiWin.Checked = True
UsWin.Checked = False
Full.Checked = False
 Frame1.Visible = False
    Frame2.Visible = True
End Sub
Private Sub HScroll1_Change()
    picCopy.Left = -HScroll1.Value
End Sub
Public Sub Capture()
    If Me.Full = True Or Me.UsWin = True Then
    Dim hdc As Long, hWnd As Long
    Dim Width As Single, Height As Single
    Dim sx As Integer, sy As Integer
    hWnd = GetForegroundWindow()
    If Me.Full.Checked Or hWnd = 0 Then ' 抓取萤幕
        hdc = GetDC(0)
        Width = Screen.Width
        Height = Screen.Height
    End If
    If Me.UsWin.Checked = True Then ' 抓取使用中的视窗
        Dim r As RECT
        hdc = GetWindowDC(hWnd)
        GetWindowRect hWnd, r
        Width = (r.Right - r.Left) * Screen.TwipsPerPixelX
        Height = (r.Bottom - r.Top) * Screen.TwipsPerPixelY
    End If
    picCopy.Width = Width
    picCopy.Height = Height
    picCopy.AutoRedraw = True
    sx = Width \ Screen.TwipsPerPixelX
    sy = Height \ Screen.TwipsPerPixelY
    BitBlt picCopy.hdc, 0, 0, sx, sy, hdc, 0, 0, vbSrcCopy
    picCopy.AutoRedraw = False

    If Full.Checked Then
        ReleaseDC 0, hdc
    Else
        ReleaseDC hWnd, hdc
    End If
    SetPicture      ' 设定 PictureBox 与卷动轴之间的关系
 End If
End Sub
Private Sub VScroll1_Change()
    picCopy.Top = -VScroll1.Value
End Sub
Private Sub Hscroll2_Change()
    picBmp.Left = -HScroll2.Value
End Sub
Private Sub Hscroll2_Scroll()
    Hscroll2_Change
End Sub
Private Sub Vscroll2_Change()
    picBmp.Top = -VScroll2.Value
End Sub
Private Sub Vscroll2_Scroll()
    Vscroll2_Change
End Sub
Private Sub picSnap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        If MtWin.Checked = True Then Me.WindowState = vbMinimized
        SetCapture picSnap.hWnd     ' 让 picSnap 得到鼠标的捕获
        tmr.Enabled = True
    End If
End Sub
Private Sub picSnap_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        tmr.Enabled = False        '
        If SnapHwnd& = 0 Then Exit Sub
        picBmp.Left = 0
        picBmp.Top = 0
        picBmp.Width = (rc.Right - rc.Left) * 15
        picBmp.Height = (rc.Bottom - rc.Top) * 15
        Dim TempDC&
        Dim newBmp&, oldBmp&
        DeskHwnd& = GetDesktopWindow()
        DeskDC& = GetWindowDC(DeskHwnd&)
        TempDC& = CreateCompatibleDC(DeskDC&)
        newBmp& = CreateCompatibleBitmap(DeskDC, _
                    rc.Right - rc.Left, rc.Bottom - rc.Top)
        oldBmp& = SelectObject(TempDC, newBmp)
        BitBlt TempDC, 0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top, _
            DeskDC, rc.Left, rc.Top, vbSrcCopy
        Me.WindowState = meStates
        BitBlt picBmp.hdc, 0, 0, _
            rc.Right - rc.Left, rc.Bottom - rc.Top, _
            TempDC, 0, 0, vbSrcCopy
        SelectObject TempDC, oldBmp
        DeleteObject newBmp: newBmp = 0
        DeleteDC TempDC
        ReleaseDC DeskHwnd, DeskDC: DeskDC = 0
        picBmp.Refresh
        ReleaseCapture      '释放鼠标的捕获
        Call Form_Resize    '
        Me.Show
        Me.SetFocus
    End If
End Sub
Private Sub tmr_Timer()
    Dim pnt As POINTAPI
    Dim newPen&, oldPen&
    DeskHwnd& = GetDesktopWindow()
    DeskDC& = GetWindowDC(DeskHwnd&)
    oldRop2& = SetROP2(DeskDC&, 10)
    GetCursorPos pnt
    SnapHwnd = WindowFromPoint(pnt.x, pnt.y)
    GetWindowRect SnapHwnd, rc
    If rc.Left < 0 Then rc.Left = 0
    If rc.Top < 0 Then rc.Top = 0
    If rc.Right > Screen.Width / 15 Then rc.Right = Screen.Width / 15
    If rc.Bottom > Screen.Height / 15 Then rc.Bottom = Screen.Height / 15
    newPen& = CreatePen(0, 3, &H0)
    oldPen& = SelectObject(DeskDC, newPen)
    Rectangle DeskDC, rc.Left, rc.Top, rc.Right, rc.Bottom
    Sleep tmr.Interval
    Rectangle DeskDC, rc.Left, rc.Top, rc.Right, rc.Bottom
    SetROP2 DeskDC, oldRop2
    SelectObject DeskDC, oldPen
    DeleteObject newPen
    ReleaseDC DeskHwnd, DeskDC: DeskDC = 0
End Sub
