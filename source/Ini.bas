Attribute VB_Name = "ModuleIni"
Option Explicit

'访问INI的函数
'用法：
'   myReadINI   读INI
'   myWriteINI  写INI
'   用法与读写注册表很类似
'               杨光宏 http://cako.126.com(VB技巧手册)



Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName As String, ByVal KeyName As String, ByVal keydefault As String, ByVal FileName As String) As Long

Public Function MyReadINI(inifile, inisection, inikey, iniDefault)

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
    Dim FileName As String
    lpDefault = Space$(254)
    lpDefault = iniDefault

    lpReturnedString = Space$(254)

    nSize = 254
    lpFileName = inifile
    lpApplicationName = inisection
    lpKeyName = inikey
    FileName = lpFileName
    retval = GetPrivateProfileString(lpApplicationName, lpKeyName, lpDefault, lpReturnedString, nSize, lpFileName)
    MyReadINI = lpReturnedString
    
End Function


Public Function myWriteINI(inifile As String, inisection As String, inikey As String, Info As String) As String
    Dim retval As Long
    retval = WritePrivateProfileString(inisection, inikey, Info, inifile)
    myWriteINI = LTrim$(Str$(retval))
End Function











Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long

Private r As Long
Private entry As String
Private iniPath As String

Function MyReadINI(AppName As String, KeyName As String, FileName As String) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   GetFromINI = Left(RetStr, GetPrivateProfileString(AppName, ByVal KeyName, "", RetStr, Len(RetStr), FileName))
End Function


Private Sub Form_Load()
    iniPath$ = App.Path + "\rwini32.ini"
    Command7_Click
End Sub
