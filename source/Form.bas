Attribute VB_Name = "modForm"
Option Explicit
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Sub FormStayOnTop(varForm As Form, ByVal OnTop As Boolean)
  Dim Handle As Long
  Dim wFlags As Long
  Dim PosFlag As Long
  Handle = varForm.hWnd
  wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
  Select Case OnTop
  Case True
    PosFlag = HWND_TOPMOST
  Case False
    PosFlag = HWND_NOTOPMOST
  End Select
  SetWindowPos Handle, PosFlag, 0, 0, 0, 0, wFlags
End Sub
Sub WaitFormClose(frm As Form)
 Do While frm Is Screen.ActiveForm
    DoEvents
  Loop
End Sub


