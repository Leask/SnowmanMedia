Attribute VB_Name = "Module1"
 Option Explicit
 Dim ERROR_SUCCESS As Integer
  Const REG_SZ = 1
    Global Const HKEY_CLASSES_ROOT = &H80000000
Declare Function OSRegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, lpdwType As Long, lpbData As Any, cbData As Long) As Long

    Declare Function OSRegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long

    Declare Function OSRegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal dwReserved As Long, ByVal fdwType As Long, lpbData As Any, ByVal cbData As Long) As Long

    Declare Function OSRegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long

  Function RegOpenKey(ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Boolean
     Dim lResult As Long
     On Error GoTo 0 ' ¹Ø±Õ´íÎóÏÝÚå
     lResult = OSRegOpenKey(hKey, lpszSubKey, phkResult)
     If lResult = 0 Then
     RegOpenKey = True
     Else
     RegOpenKey = False
     End If
    End Function
    Function RegSetStringValue(ByVal hKey As Long, ByVal strValueName As String, ByVal strData As String, Optional ByVal fLog) As Boolean
     Dim lResult As Long
     On Error GoTo 0
     lResult = OSRegSetValueEx(hKey, strValueName, 0&, REG_SZ, ByVal strData, LenB(StrConv(strData, vbFromUnicode)) + 1)
     If lResult = 0 Then
     RegSetStringValue = True
     Else
     RegSetStringValue = False
     End If
    End Function
    Function StripTerminator(ByVal strString As String) As String
     Dim intZeroPos As Integer
     intZeroPos = InStr(strString, Chr$(0))
If intZeroPos > 0 Then
    StripTerminator = Left$(strString, intZeroPos - 1)
     Else
     StripTerminator = strString
     End If
    End Function
    Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String, strData As String) As Boolean
     Dim lResult As Long
     Dim lValueType As Long
     Dim strBuf As String
     Dim lDataBufSize As Long
     RegQueryStringValue = False
     On Error GoTo 0
     lResult = OSRegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
     If lResult = ERROR_SUCCESS Then
     If lValueType = REG_SZ Then
     strBuf = String(lDataBufSize, " ")
     lResult = OSRegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
     If lResult = ERROR_SUCCESS Then
     RegQueryStringValue = True
     strData = StripTerminator(strBuf)
     End If
     End If
     End If
    End Function
