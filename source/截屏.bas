Attribute VB_Name = "Module1032025"
Option Explicit


Public Const GWL_WNDPROC = (-4)
Public Const WM_USERB = &H400
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONUP = &H202

Public Const NIM_ADD = 0
Public Const NIM_MODIFY = 1
Public Const NIM_DELETE = 2

Public Const NIF_MESSAGE = 1
Public Const NIF_ICON = 2
Public Const NIF_TIP = 4

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Type NOTIFYICONDATA
       cbSize As Long
       hWnd As Long
       uID As Long
       uFlags As Long
       uCallbackMessage As Long
       hIcon As Long
       szTip As String * 64
End Type

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function SetKeyboardHook Lib "KeybHook" (ByVal hwndPost As Long, ByVal Msg As Long) As Long
Declare Function ReleaseKeyboardHook Lib "KeybHook" () As Long

Public prevWndProc As Long

Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_USERB Then
        If wParam = vbKeyF7 And (lParam And &H80000000) <> 0 Then
            Formjp.Capture
            Formjp.picCopy.Refresh
            Formjp.WindowState = vbNormal
            Formjp.Show
        End If
    ElseIf Msg = WM_USERB + 100 Then
        If lParam = WM_LBUTTONDBLCLK Then
            Formjp.WindowState = vbNormal
            Formjp.Show
        ElseIf lParam = WM_LBUTTONUP Then
            Formjp.PopupMenu Formjp.mDx
        End If
    End If
    WndProc = CallWindowProc(prevWndProc, hWnd, Msg, wParam, lParam)
End Function

