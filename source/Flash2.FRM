VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7260
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu a 
      Caption         =   "文件(&F)"
      Begin VB.Menu b 
         Caption         =   "打开(&O)..."
      End
      Begin VB.Menu c 
         Caption         =   "-"
      End
      Begin VB.Menu d 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu e 
      Caption         =   "窗口(&W)"
   End
   Begin VB.Menu f 
      Caption         =   "帮助(&H)"
   End
   Begin VB.Menu g 
      Caption         =   "关于(&A)"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const GW_HWNDNEXT = 2

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private old_parent As Long
Private child_hwnd As Long

' Return the window handle for an instance handle.
Private Function InstanceToWnd(ByVal target_pid As Long) As Long
Dim test_hwnd As Long
Dim test_pid As Long
Dim test_thread_id As Long

    ' Get the first window handle.
    test_hwnd = FindWindow(ByVal 0&, ByVal 0&)

    ' Loop until we find the target or we run out
    ' of windows.
    Do While test_hwnd <> 0
        ' See if this window has a parent. If not,
        ' it is a top-level window.
        If GetParent(test_hwnd) = 0 Then
            ' This is a top-level window. See if
            ' it has the target instance handle.
            test_thread_id = GetWindowThreadProcessId(test_hwnd, test_pid)

            If test_pid = target_pid Then
                ' This is the target.
                InstanceToWnd = test_hwnd
                Exit Do
            End If
        End If

        ' Examine the next window.
        test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    Loop
End Function

Private Sub cmdFree_Click()
    SetParent child_hwnd, old_parent

    cmdRun.Enabled = True
    cmdFree.Enabled = False
End Sub



Private Sub b_Click()
Dim pid As Long
Dim buf As String
Dim buf_len As Long
Dim styles As Long

    ' Start the program.
    pid = Shell(App.Path + "\a.exe", vbNormalFocus)
    If pid = 0 Then
        MsgBox "Error starting program"
        Exit Sub
    End If

    ' Get the window handle.
    child_hwnd = InstanceToWnd(pid)

    ' Reparent the program so it lies inside
    ' the PictureBox.
    old_parent = SetParent(child_hwnd, MDIForm1.hwnd)


End Sub

Private Sub MDIForm_Load()
Dim pid As Long
Dim buf As String
Dim buf_len As Long
Dim styles As Long

    ' Start the program.
    pid = Shell(App.Path + "\a.exe", vbNormalFocus)
    If pid = 0 Then
        MsgBox "Error starting program"
        Exit Sub
    End If

    ' Get the window handle.
    child_hwnd = InstanceToWnd(pid)

    ' Reparent the program so it lies inside
    ' the PictureBox.
    old_parent = SetParent(child_hwnd, MDIForm1.hwnd)


End Sub
