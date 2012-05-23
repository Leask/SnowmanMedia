Attribute VB_Name = "basInitEntry"
Option Explicit
Private Type FormPosition
    Left    As Long
    Top     As Long
    Width   As Long
    Height  As Long
    Maxed   As Boolean
End Type
Private sDefInitFileName As String
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Sub AddRecentFile(ByVal sNewFileName As String, mnuRecent As Variant, Optional ByVal iMaxEntries As Integer = 8, Optional ByVal iMaxFileNameLen As Integer = 60)
Dim lRet        As Long
Dim iArrayCnt   As Integer
Dim iFileCnt    As Integer
Dim sFilename   As String
Dim saFiles()    As String
    ReDim saFiles(iMaxEntries)
    saFiles(0) = sNewFileName
    iFileCnt = 1
    sFilename = GetInitEntryB("Recent Files", "File " & CStr(iFileCnt), "")
    While Len(sFilename) > 0 And iArrayCnt < iMaxEntries
        If LCase$(sFilename) <> LCase$(sNewFileName) Then
            iArrayCnt = iArrayCnt + 1
            saFiles(iArrayCnt) = sFilename
        End If
        iFileCnt = iFileCnt + 1
        sFilename = GetInitEntryB("Recent Files", "File " & CStr(iFileCnt), "")
    Wend
    ReDim Preserve saFiles(iArrayCnt)
    lRet = SetInitEntryB("Recent Files")
    For iFileCnt = 0 To iArrayCnt
        lRet = SetInitEntryB("Recent Files", "File " & CStr(iFileCnt + 1), saFiles(iFileCnt))
    Next iFileCnt
    Call GetRecentFiles(mnuRecent, iMaxEntries, iMaxFileNameLen)
    mnuRecent(0).Checked = (mnuRecent(0).Caption <> "(Empty)")
End Sub
Public Sub GetRecentFiles(mnuRecent As Variant, Optional ByVal iMaxEntries As Integer = 8, Optional ByVal iMaxFileNameLen As Integer = 60)
Dim iIdx        As Integer
Dim iFileCnt    As Integer
Dim iFullCnt    As Integer
Dim iMenuCnt    As Integer
Dim sFilename   As String
    On Error GoTo LocalError
    iMenuCnt = mnuRecent.UBound
    For iIdx = 1 To iMenuCnt
        Unload mnuRecent(iIdx)
    Next iIdx
    mnuRecent(0).Checked = False
    mnuRecent(0).Tag = ""
    mnuRecent(0).Enabled = False
    mnuRecent(0).Caption = "(Empty)"
    sFilename = GetInitEntryB("Recent Files", "File " & CStr(iFullCnt + 1), "")
    While Len(sFilename) > 0 And iFileCnt <= iMaxEntries
        If Exists(sFilename) Then
            If iFileCnt > 0 Then
                Load mnuRecent(iFileCnt)
            End If
            mnuRecent(iFileCnt).Caption = "&" & CStr(iFileCnt + 1) & " " & _
                ShortenFileName(sFilename, iMaxFileNameLen)
            mnuRecent(iFileCnt).Tag = sFilename
            mnuRecent(iFileCnt).Enabled = True
            mnuRecent(iFileCnt).Visible = True
            iFileCnt = iFileCnt + 1
        End If
        iFullCnt = iFullCnt + 1
        sFilename = GetInitEntryB("Recent Files", "File " & CStr(iFullCnt + 1), "")
    Wend
NormalExit:
    Exit Sub
LocalError:
    MsgBox Err.Description, vbExclamation, App.EXEName
    Resume NormalExit
End Sub
Private Function Exists(ByVal sFilename As String) As Boolean
    If Len(Trim$(sFilename)) > 0 Then
        On Error Resume Next
        sFilename = Dir$(sFilename)
        Exists = Err.Number = 0 And Len(sFilename) > 0
    Else
        Exists = False
    End If
End Function
Public Sub RemoveRecentFile(ByVal sRemoveFileName As String, mnuRecent As Variant, Optional ByVal iMaxEntries As Integer = 8, Optional ByVal iMaxFileNameLen As Integer = 60)
Dim lRet        As Long
Dim iArrayCnt   As Integer
Dim iFileCnt    As Integer
Dim sFilename   As String
Dim saFiles()    As String
    ReDim saFiles(iMaxEntries)
    iFileCnt = 1
    sFilename = GetInitEntryB("Recent Files", "File " & CStr(iFileCnt), "")
    While Len(sFilename) > 0 And iArrayCnt < iMaxEntries
        If LCase$(sFilename) <> LCase$(sRemoveFileName) Then
            saFiles(iArrayCnt) = sFilename
            iArrayCnt = iArrayCnt + 1
        End If
        iFileCnt = iFileCnt + 1
        sFilename = GetInitEntryB("Recent Files", "File " & CStr(iFileCnt), "")
    Wend
    ReDim Preserve saFiles(iArrayCnt - 1)
    lRet = SetInitEntryB("Recent Files")
    For iFileCnt = 0 To iArrayCnt - 1
        lRet = SetInitEntryB("Recent Files", "File " & CStr(iFileCnt + 1), saFiles(iFileCnt))
    Next iFileCnt
    Call GetRecentFiles(mnuRecent, iMaxEntries, iMaxFileNameLen)
End Sub
Private Function ShortenFileName(ByVal sFilename As String, ByVal iMaxLen As Integer) As String
Dim iLen        As Integer
Dim iSlashPos   As Integer
    On Error GoTo LocalError
    If Len(sFilename) > iMaxLen Then
        iLen = iMaxLen - 3
        iSlashPos = InStr(sFilename, "\")
        While (iSlashPos > 0) And (Len(sFilename) > iLen)
            sFilename = Mid$(sFilename, iSlashPos)
            iSlashPos = InStr(2, sFilename, "\")
        Wend
        If Len(sFilename) > iLen Then
            sFilename = "..." & Mid$(sFilename, Len(sFilename) - iLen + 1)
        Else
            sFilename = "..." & sFilename
        End If
    End If
    ShortenFileName = sFilename
NormalExit:
    Exit Function
LocalError:
    MsgBox Err.Description, vbExclamation, App.EXEName
    Resume NormalExit
End Function
Public Function GetInitEntryB(ByVal sSection As String, ByVal sKeyName As String, Optional ByVal sDefault As String = "", Optional ByVal sInitFileName As String = "") As String
Dim sBuffer As String
Dim sInitFile As String
    If Len(sInitFileName) = 0 Then
        If Len(sDefInitFileName) = 0 Then
            sDefInitFileName = App.Path
            If Right$(sDefInitFileName, 1) <> "\" Then
                sDefInitFileName = sDefInitFileName & "\"
            End If
            sDefInitFileName = sDefInitFileName & App.EXEName & ".ini"
        End If
        sInitFile = sDefInitFileName
    Else
        sInitFile = sInitFileName
    End If
    sBuffer = String$(2048, " ")
    GetInitEntryB = Left$(sBuffer, GetPrivateProfileString(sSection, ByVal sKeyName, sDefault, sBuffer, Len(sBuffer), sInitFile))
End Function
Public Sub SaveFormSize(frmForm As Form, Optional ByVal sInitFileName As String = "")
Dim lRet        As Long
Dim sData       As String
Dim saSizes()   As String
    ReDim saSizes(4)
    If frmForm.WindowState = vbNormal Then
        saSizes(0) = CStr(frmForm.Left)
        saSizes(1) = CStr(frmForm.Top)
        saSizes(2) = CStr(frmForm.Width)
        saSizes(3) = CStr(frmForm.Height)
        saSizes(4) = "False"
    Else
        sData = GetInitEntryB("Positions", frmForm.Name, "", sInitFileName)
        If Len(sData) = 0 Then
            saSizes(0) = "-1"
            saSizes(1) = "-1"
            saSizes(2) = "-1"
            saSizes(3) = "-1"
        Else
            saSizes() = Split(sData, ",")
            ReDim Preserve saSizes(4)
        End If
        saSizes(4) = CStr(frmForm.WindowState = vbMaximized)
    End If
    lRet = SetInitEntryB("Positions", frmForm.Name, Join(saSizes, ","), sInitFileName)
End Sub
Public Sub RestoreFormSize(frmForm As Form, Optional ByVal sInitFileName As String = "")
Dim sData       As String
Dim saSizes()   As String
Dim uPosition   As FormPosition
    With uPosition
        sData = GetInitEntryB("Positions", frmForm.Name, "", sInitFileName)
        If Len(sData) = 0 Then
            .Left = frmForm.Left
            .Top = frmForm.Top
            .Width = frmForm.Width
            .Height = frmForm.Height
            .Maxed = frmForm.WindowState = vbMaximized
        Else
            saSizes() = Split(sData, ",")
            If UBound(saSizes) < 4 Then
                ReDim Preserve saSizes(4)
            End If
            .Left = Val(Trim$(saSizes(0)))
            .Top = Val(Trim$(saSizes(1)))
            .Width = Val(Trim$(saSizes(2)))
            .Height = Val(Trim$(saSizes(3)))
            .Maxed = LCase$(Trim$(saSizes(4))) = "true"
        End If
        If .Width < 150 Then
            .Width = frmForm.Width
        ElseIf .Width > Screen.Width Then
            .Width = Screen.Width
        End If
        If .Left < 0 Then
            .Left = frmForm.Left
        End If
        If .Left > Screen.Width - .Width Then
            .Left = Screen.Width - .Width
        End If
        If .Height < 150 Then
            .Height = frmForm.Height
        ElseIf .Height > Screen.Height Then
            .Height = Screen.Height
        End If
        If .Top < 0 Then
            .Top = frmForm.Top
        End If
        If .Top > Screen.Height - .Height Then
            .Top = Screen.Height - .Height
        End If
        frmForm.Move .Left, .Top, .Width, .Height
        If .Maxed Then
            frmForm.WindowState = vbMaximized
        End If
    End With
End Sub
Public Function SetInitEntryB(ByVal sSection As String, Optional ByVal sKeyName As String, Optional ByVal sValue As String, Optional ByVal sInitFileName As String = "") As Long
Dim sInitFile As String
    If Len(sInitFileName) = 0 Then
        If Len(sDefInitFileName) = 0 Then
            sDefInitFileName = App.Path
            If Right$(sDefInitFileName, 1) <> "\" Then
                sDefInitFileName = sDefInitFileName & "\"
            End If
            sDefInitFileName = sDefInitFileName & App.EXEName & ".ini"
        End If
        sInitFile = sDefInitFileName
    Else
        sInitFile = sInitFileName
    End If
    If Len(sKeyName) > 0 And Len(sValue) > 0 Then
        SetInitEntryB = WritePrivateProfileString(sSection, ByVal sKeyName, ByVal sValue, sInitFile)
    ElseIf Len(sKeyName) > 0 Then
        SetInitEntryB = WritePrivateProfileString(sSection, ByVal sKeyName, vbNullString, sInitFile)
    Else
        SetInitEntryB = WritePrivateProfileString(sSection, vbNullString, vbNullString, sInitFile)
    End If
End Function
