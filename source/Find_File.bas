Attribute VB_Name = "Find_File"
'Made by Michael Kruse
Option Explicit
Public Const MAX_PATH = 260
Public Const UnicodeTypeLib = True
Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
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
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal sDrive As String) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long


Public Function FindFiles(sTarget As String, Optional _
                   ByVal Start As String) As Collection
    
Dim ab() As Byte
Static TypeDrev As String
Dim hFiles As Long, f As Boolean
Static sName As String, sSpec As String, nFound As New Collection
Static fd As WIN32_FIND_DATA, iLevel As Long
Dim sEmpty, INVALID_HANDLE_VALUE
      
If Start = sEmpty Then Start = CurDir$
   'Maintain level to ensure collection is cleared first time
    If iLevel = 0 Then
        Set nFound = Nothing
        Start = NormalizePath(Start)
    End If
    iLevel = iLevel + 1
   'Find first file (get handle to find)
    hFiles = FindFirstFile(Start & "*.*", fd)
    f = (hFiles <> INVALID_HANDLE_VALUE)
    Do While f
        ab = fd.cFileName
        sName = ByteZToStr(ab)
       'Skip . and ..
        If Left$(sName, 1) <> "." Then
            sSpec = Start & sName
            If fd.dwFileAttributes And vbDirectory Then
               'Call recursively on each directory
                DoEvents
                FindFiles sTarget, sSpec & "\"
            Else
                If InStr(sTarget, "*") > 0 Then
                    If StrComp(Right$(sName, 3), Right$(sTarget, 3), 1) = 0 Then ' Text comparison
                   'Store found files in collection
                    nFound.Add sSpec
                ElseIf StrComp(sName, sTarget, 1) = 0 Then ' Text comparison
                   'Store found files in collection
                    nFound.Add sSpec
                End If
            End If
        End If
End If
   'Keep looping until no more files
    f = FindNextFile(hFiles, fd)
    Loop
f = FindClose(hFiles)
'Return the matching files in collection
Set FindFiles = nFound
iLevel = iLevel - 1

End Function
Function ByteZToStr(ab() As Byte) As String
    
    If UnicodeTypeLib Then
        ByteZToStr = ab
    Else
        ByteZToStr = StrConv(ab, vbUnicode)
    End If
    ByteZToStr = Left$(ByteZToStr, lstrlen(ByteZToStr))
End Function

Function NormalizePath(sPath As String) As String
    If Right$(sPath, 1) <> "\" Then
        NormalizePath = sPath & "\"
    Else
        NormalizePath = sPath
    End If
End Function

