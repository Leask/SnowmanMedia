Attribute VB_Name = "Module1"
Option Explicit
Public Const MAX_PATH = 260
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public Declare Function FindFirstFile Lib "kernel32" Alias _
        "FindFirstFileA" (ByVal lpFileName As String, _
        lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias _
        "FindNextFileA" (ByVal hFindFile As Long, _
        lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal _
        hFindFile As Long) As Long
Public Declare Function SetCurrentDirectory Lib "kernel32" Alias _
        "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

Sub AllSearch(sPath As String, sFile As String)
    Dim xf As WIN32_FIND_DATA
    Dim ff As WIN32_FIND_DATA
    Dim findhandle As Long
    Dim lFindFile As Long
    Dim astr As String
    Dim bstr As String
    
    lFindFile = FindFirstFile(sPath + "\" + sFile, ff)
    'Debug.Print sPath + "\" + sFile
    If lFindFile > 0 Then
        Do
            Form1.List1.AddItem ff.cFileName
        Loop Until (FindNextFile(lFindFile, ff) = 0)
        FindClose lFindFile
    End If
    'Debug.Print Form1.List1.ListCount
    
    astr = sPath + "\" + "*.*"
    findhandle = FindFirstFile(astr, xf)
    DoEvents
    Do
        If xf.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Then
            If Asc(xf.cFileName) <> Asc(".") Then
                bstr = sPath + "\" + Left$(xf.cFileName, InStr(xf.cFileName, Chr(0)) - 1)
                'Debug.Print bstr
                AllSearch bstr, sFile
                
                'lFindFile = FindFirstFile(bstr, ff)
                'Debug.Print bstr
                'Do
                '    Form1.List1.AddItem ff.cFileName
                    'Debug.Print ff.cFileName
                '    ff.cFileName = ""
                'Loop Until (FindNextFile(lFindFile, ff) = 0)
                'FindClose lFindFile
            End If
        End If
        xf.cFileName = ""
    Loop Until (FindNextFile(findhandle, xf) = 0)
    FindClose findfile
End Sub
