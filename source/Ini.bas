Attribute VB_Name = "ModuleIni"
Option Explicit



Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName As String, ByVal KeyName As String, ByVal keydefault As String, ByVal Filename As String) As Long

Public Function myReadINI(inifile, inisection, inikey, iniDefault)

'Fail fracefully if no file / wrong file is specified.
'If no section (appname), default is first appname
'if no key, default is first key


    Dim lpApplicationName As String
    Dim lpKeyName As String
    Dim lpDefault As String
    Dim lpReturnedString As String
    Dim nSize As Long
    Dim lpFileName As String
    Dim retval As Long
    Dim Filename As String
    lpDefault = Space$(254)
    lpDefault = iniDefault

    lpReturnedString = Space$(254)

    nSize = 254
    lpFileName = inifile
    lpApplicationName = inisection
    lpKeyName = inikey
    Filename = lpFileName
    retval = GetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, lpReturnedString, nSize, lpFileName)
    myReadINI = lpReturnedString
    
End Function


Public Function myWriteINI(inifile As String, inisection As String, inikey As String, Info As String) As String
    Dim retval As Long
    retval = WritePrivateProfileString(inisection, inikey, Info, inifile)
    myWriteINI = LTrim(Str$(retval))
End Function



