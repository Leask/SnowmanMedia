VERSION 5.00
Begin VB.Form frmMainB 
   Caption         =   "Snowman Media Screen Zoom  1.0"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3255
   Icon            =   "放大镜.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   217
   Begin VB.HScrollBar hsbZoom 
      Height          =   240
      LargeChange     =   10
      Left            =   450
      Max             =   1000
      Min             =   25
      TabIndex        =   1
      Top             =   0
      Value           =   25
      Width           =   1230
   End
   Begin VB.TextBox txtZoom 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Height          =   240
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "1000%"
      Top             =   0
      Width           =   570
   End
   Begin VB.CheckBox chkOnTop 
      BackColor       =   &H0000FFFF&
      DownPicture     =   "放大镜.frx":1582
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   225
      Picture         =   "放大镜.frx":1678
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   240
   End
   Begin VB.CheckBox chkGrid 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   0
      Picture         =   "放大镜.frx":176E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox picZoom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   30
      ScaleHeight     =   107
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   315
      Width           =   1605
   End
   Begin VB.Timer tmrZoom 
      Interval        =   50
      Left            =   1710
      Top             =   360
   End
End
Attribute VB_Name = "frmMainB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    x   As Long
    y   As Long
End Type
Private Type SizeRect
    Left    As Long
    Top     As Long
    Width   As Long
    Height  As Long
End Type
Private Type RectAPI
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type
Private Const SRCCOPY           As Long = &HCC0020
Private Const PATCOPY           As Long = &HF00021
Private Const SWP_NOMOVE        As Long = 2
Private Const SWP_NOSIZE        As Long = 1
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_FLAGS         As Long = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
Private Const HWND_TOPMOST      As Long = -1
Private Const HWND_NOTOPMOST    As Long = -2
Private mfScale As Single
Private mlOldX  As Long
Private mlOldY  As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RectAPI) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Function CreateCheckeredBrush(ByVal hdc As Long, ByVal lColor1 As Long, ByVal lColor2 As Long) As Long
Dim x           As Long
Dim y           As Long
Dim lRet        As Long
Dim hBitmapDC   As Long
Dim hBitmap     As Long
Dim hOldBitmap  As Long
    If lColor1 < 0 Then
        lColor1 = GetSysColor(lColor1 And &HFF&)
    End If
    If lColor2 < 0 Then
        lColor2 = GetSysColor(lColor2 And &HFF&)
    End If
    hBitmapDC = CreateCompatibleDC(hdc)
    hBitmap = CreateCompatibleBitmap(hdc, 8, 8)
    hOldBitmap = SelectObject(hBitmapDC, hBitmap)
    For y = 0 To 6 Step 2
        For x = 0 To 6 Step 2
            lRet = SetPixelV(hBitmapDC, x, y, lColor1)
            lRet = SetPixelV(hBitmapDC, x + 1, y, lColor2)
            lRet = SetPixelV(hBitmapDC, x, y + 1, lColor2)
            lRet = SetPixelV(hBitmapDC, x + 1, y + 1, lColor1)
        Next x
    Next y
    hBitmap = SelectObject(hBitmapDC, hOldBitmap)
    CreateCheckeredBrush = CreatePatternBrush(hBitmap)
    lRet = DeleteDC(hBitmapDC)
    lRet = DeleteObject(hBitmap)
End Function
Private Sub DoZoom(ptMouse As POINTAPI)
Dim lRet        As Long
Dim lTemp       As Long
Dim hWndDesk    As Long
Dim hDCDesk     As Long
Dim sizSrce     As SizeRect
Dim sizDest     As SizeRect
    hWndDesk = GetDesktopWindow()
    hDCDesk = GetDC(hWndDesk)
    With sizDest
        .Left = 0
        .Top = 0
        .Width = picZoom.ScaleWidth
        .Height = picZoom.ScaleHeight
    End With
    With sizSrce
        .Left = ptMouse.x - Int((sizDest.Width / 2) / mfScale)
        .Top = ptMouse.y - Int((sizDest.Height / 2) / mfScale)
        .Width = Int(sizDest.Width / mfScale)
        .Height = Int(sizDest.Height / mfScale)
        lTemp = Int(.Width * mfScale)
        If lTemp > sizDest.Width Then
            sizDest.Width = lTemp
        ElseIf lTemp < sizDest.Width Then
            .Width = .Width + 1
            sizDest.Width = lTemp + mfScale
        End If
        lTemp = Int(.Height * mfScale)
        If lTemp > sizDest.Height Then
            sizDest.Height = lTemp
        ElseIf lTemp < sizDest.Height Then
            .Height = .Height + 1
            sizDest.Height = lTemp + mfScale
        End If
    End With
    picZoom.Cls
    lRet = StretchBlt(picZoom.hdc, sizDest.Left, sizDest.Top, sizDest.Width, sizDest.Height, hDCDesk, sizSrce.Left, sizSrce.Top, sizSrce.Width, sizSrce.Height, SRCCOPY)
    lRet = ReleaseDC(hWndDesk, hDCDesk)
    If chkGrid.Value = vbChecked Then
        Call DrawGrid
    End If
    picZoom.Refresh
End Sub
Private Sub DrawGrid()
Dim iWidth      As Integer
Dim iHeight     As Integer
Dim lRet        As Long
Dim hBrush      As Long
Dim hOldBrush   As Long
Dim fX          As Single
Dim fY          As Single
    If mfScale >= 3 Then
        hBrush = CreateCheckeredBrush(picZoom.hdc, &H808080, &HC0C0C0)
        hOldBrush = SelectObject(picZoom.hdc, hBrush)
        iWidth = picZoom.ScaleWidth
        iHeight = picZoom.ScaleHeight
        For fX = 0 To iWidth Step mfScale
            lRet = PatBlt(picZoom.hdc, Int(fX), 0, 1, iHeight, PATCOPY)
        Next
        For fY = 0 To iHeight Step mfScale
            lRet = PatBlt(picZoom.hdc, 0, Int(fY), iWidth, 1, PATCOPY)
        Next
        hBrush = SelectObject(picZoom.hdc, hOldBrush)
        lRet = DeleteObject(hBrush)
    End If
End Sub
Private Function ValidScale(ByVal fScale As Single) As Single
    If fScale * 100 > hsbZoom.Max Then
        fScale = hsbZoom.Max / 100
    ElseIf fScale * 100 < hsbZoom.Min Then
        fScale = hsbZoom.Min / 100
    End If
    ValidScale = fScale
End Function
Private Sub LoadSettings()
    Call RestoreFormSize(Me)
    hsbZoom.Value = GetInitEntryB("Settings", "Zoom", CStr(200))
    hsbZoom_Change
    chkGrid.Value = IIf(LCase$(GetInitEntryB("Settings", "Grid", "False")) = "true", vbChecked, vbUnchecked)
    chkGrid_Click
    chkOnTop.Value = IIf(LCase$(GetInitEntryB("Settings", "OnTop", "False")) = "true", vbChecked, vbUnchecked)
    chkOnTop_Click
End Sub
Private Sub SaveSettings()
Dim lRet As Long
    Call SaveFormSize(Me)
    lRet = SetInitEntryB("Settings", "Zoom", hsbZoom.Value)
    lRet = SetInitEntryB("Settings", "Grid", CStr(chkGrid.Value = vbChecked))
    lRet = SetInitEntryB("Settings", "OnTop", CStr(chkOnTop.Value = vbChecked))
End Sub
Private Sub chkGrid_Click()
    mlOldX = -100
    If picZoom.Visible Then
        picZoom.SetFocus
    End If
End Sub
Private Sub chkOnTop_Click()
Dim lRet    As Long
Dim lWinPos As Long
    lWinPos = IIf(chkOnTop.Value = vbChecked, HWND_TOPMOST, HWND_NOTOPMOST)
    lRet = SetWindowPos(Me.hWnd, lWinPos, 0, 0, 0, 0, SWP_FLAGS)
    If picZoom.Visible Then
        picZoom.SetFocus
    End If
End Sub
Private Sub Form_Load()
    Call LoadSettings
End Sub
Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        If Me.Width < 1680 Then
            Me.Width = 1680
        ElseIf Me.Height < 1680 Then
            Me.Height = 1680
        Else
            chkGrid.Move 0, 0
            chkOnTop.Move chkGrid.Width, 0
            hsbZoom.Move chkGrid.Width + chkOnTop.Width, 0, Me.ScaleWidth - txtZoom.Width - chkGrid.Width - chkOnTop.Width
            txtZoom.Move Me.ScaleWidth - txtZoom.Width, -1
            picZoom.Move 0, hsbZoom.Height, Me.ScaleWidth, Me.ScaleHeight - hsbZoom.Height
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call SaveSettings
End Sub
Private Sub hsbZoom_Change()
    txtZoom.Text = Format$(hsbZoom.Value / 100, "####%")
    mfScale = CSng(hsbZoom.Value) / 100!
    If picZoom.Visible Then
        picZoom.SetFocus
    End If
    mlOldX = -100
End Sub
Private Sub hsbZoom_Scroll()
    hsbZoom_Change
End Sub
Private Sub tmrZoom_Timer()
Dim lRet    As Long
Dim ptMouse As POINTAPI
Static lElapsed As Long
    If Me.WindowState <> vbMinimized Then
        lElapsed = lElapsed + tmrZoom.Interval
        lRet = GetCursorPos(ptMouse)
        With ptMouse
            If (.x <> mlOldX) Or (.y <> mlOldY) Or (lElapsed >= 250) Then
                Call DoZoom(ptMouse)
                If lElapsed >= 250 Then
                    If chkOnTop.Value = vbChecked Then
                        lRet = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS)
                    End If
                End If
                lElapsed = 0
            End If
            mlOldX = .x
            mlOldY = .y
        End With
    End If
End Sub
Private Sub txtZoom_GotFocus()
    With txtZoom
        .Text = CStr(Val(.Text))
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtZoom_KeyPress(KeyAscii As Integer)
    If KeyAscii > 31 And (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then
        Beep
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then
        picZoom.SetFocus
        DoEvents
        txtZoom.SetFocus
        KeyAscii = 0
    End If
End Sub
Private Sub txtZoom_LostFocus()
    mfScale = ValidScale(Val(txtZoom.Text) / 100)
    hsbZoom.Value = mfScale * 100
    txtZoom.Text = Format$(mfScale, "####%")
End Sub
