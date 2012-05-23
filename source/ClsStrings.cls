VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ToUpperCase(pStr As String) As String

On Error GoTo ToUpperCaseERR

    ToUpperCase = UCase$(pStr)

ToUpperCaseEXIT:
    Exit Function
ToUpperCaseERR:
    ToUpperCase = Null
    Err.Raise (Err.Number)
    Resume ToUpperCaseEXIT
End Function

Public Function ToLowerCase(pStr As String) As String

On Error GoTo ToLowerCaseERR

    ToLowerCase = LCase$(pStr)
    
ToLowerCaseEXIT:
    Exit Function
ToLowerCaseERR:
    ToLowerCase = Null
    Err.Raise (Err.Number)
    Resume ToLowerCaseEXIT
End Function

Public Function Proper(pStr As String) As String
    
'Converts the first letter in the word to uppercase and the
'rest of the characters to lower case

On Error GoTo ProperERR

    Proper = StrConv(pStr, vbProperCase)

ProperEXIT:
    Exit Function
ProperERR:
    Proper = Null
    Err.Raise (Err.Number)
    Resume ProperEXIT
End Function

Public Function ReplaceString(pStr As String, strFrom As String, strTo As String, Optional ignoreCase As Variant) As String

'Replaces all occurences of strFrom with strTo in pStr
'ignoreCase tells us to make case insensitive comparisons

On Error GoTo ReplaceStringERR

Dim opStr As String
Dim s As New clsStringAnalyser
Dim res As Variant
Dim c As New Collection
    
    If IsMissing(ignoreCase) Then
        ignoreCase = True
    End If
    
    c.Add strFrom
    
    res = s.PopWord(pStr, ignoreCase, c)
    
    Do While Not IsNull(res)
        opStr = opStr + res + strTo
        res = s.PopWord()
    Loop
    
    opStr = Left(opStr, Len(opStr) - Len(strTo))
    
    ReplaceString = opStr
    
ReplaceStringEXIT:
    Exit Function
ReplaceStringERR:
    ReplaceString = Null
    Err.Raise (Err.Number)
    Resume ReplaceStringEXIT
End Function

Public Function RemoveString(pStr As String, strRemove As String, Optional ignoreCase As Variant) As String

'Removes all occurences of strRemove from pStr
'ignoreCase tells us to ignore the case when making comparisons

On Error GoTo RemoveStringERR

Dim opStr As String
Dim c As New Collection
Dim s As New clsStringAnalyser
Dim res As Variant

    If IsMissing(ignoreCase) Then
        ignoreCase = True
    End If
    
    c.Add strRemove
    
    res = s.PopWord(pStr, ignoreCase, c)
    
    Do While Not IsNull(res)
        opStr = opStr + res
        res = s.PopWord()
    Loop
    
    RemoveString = opStr
    
RemoveStringEXIT:
    Exit Function
RemoveStringERR:
    RemoveString = Null
    Err.Raise (Err.Number)
    Resume RemoveStringEXIT
End Function

Public Function ToGgLeCaSe(pStr As String) As String

'Toggles each character between cases

On Error GoTo ToGgLeCaSeERR

Dim i As Long
Dim m As Long
Dim c As String * 1
Dim oppStr As String

    m = Len(pStr)
    
    For i = 1 To m
        c = Mid(pStr, i, 1)
        If c = UCase(c) Then
            oppStr = oppStr + LCase(c)
        Else
            oppStr = oppStr + UCase(c)
        End If
    Next

    ToGgLeCaSe = oppStr
    
ToGgLeCaSeEXIT:
    Exit Function
ToGgLeCaSeERR:
    ToGgLeCaSe = Null
    Err.Raise (Err.Number)
    Resume ToGgLeCaSeEXIT
End Function

Public Function IsDateDDMMYYYY(ByRef pStr As String, _
                                Optional appendCentury As Variant, _
                                Optional delimiters As Variant) As Variant
    
'Input
'   pStr = The string to be considered
'   appendCentury = True = Append the current century when the year part of the string is
'                          between 1 and 99
'                   False = Do not append the century.  Forces entry of the century
'   delimiters = A collection of strings used for delimiting the date parts.
'                    May be any length and mixed within the input string
'Result
'   NULL = The date is not valid
'   <Value> = The string has been validated as a DDMMYYYY date and is returned formatted as
'               DD/MM/YYYY

On Error GoTo IsDateDDMMYYYYERR

Dim dateDelimiters As New Collection
Dim dateParts As New Collection
Dim s As New clsStringAnalyser
Dim res As Variant
Dim appCentury As Boolean
Dim sPossibleDate As String

    If IsMissing(delimiters) Then
        dateDelimiters.Add "/"
    Else
        Set dateDelimiters = delimiters
    End If
    
    If IsMissing(appendCentury) Then
        appCentury = False
    Else
        appCentury = appendCentury
    End If
        
    'Determine the date parts
    res = s.PopWord(pStr, True, dateDelimiters)

    Do While Not IsNull(res)
        dateParts.Add res
        res = s.PopWord()
    Loop
    
    'we need three parts to the date
    If dateParts.Count <> 3 Then
        IsDateDDMMYYYY = Null
        Exit Function
    End If
    
    'Day is invalid
    If Val(dateParts(1)) < 1 Or Val(dateParts(1)) > 31 Then
        IsDateDDMMYYYY = Null
        Exit Function
    End If
    
    'Month is invalid
    If Val(dateParts(2)) < 1 Or Val(dateParts(2)) > 12 Then
        IsDateDDMMYYYY = Null
        Exit Function
    End If
    
    'Got to enter at least two digits in the century
    If Len(dateParts(3)) < 2 Then
        IsDateDDMMYYYY = Null
        Exit Function
    End If
    
    'append the century if we are instructed to do so and
    'the century part of the date is two digits
    If appCentury And Val(dateParts(3)) >= 0 And Val(dateParts(3)) < 100 Then
        dateParts.Add Str(dateParts(3) + (Val(Left(Format$(Now, "yyyy"), 2)) * 100))
        dateParts.Remove (3)
    End If
    
    If dateParts(3) < 1000 Or dateParts(3) > 9999 Then
        IsDateDDMMYYYY = Null
        Exit Function
    End If
    
    sPossibleDate = dateParts(1) + "/" + dateParts(2) + "/" + dateParts(3)
    
    If IsDate(sPossibleDate) Then
        IsDateDDMMYYYY = sPossibleDate
    Else
        IsDateDDMMYYYY = Null
    End If
    
IsDateDDMMYYYYEXIT:
    Exit Function
IsDateDDMMYYYYERR:
    IsDateDDMMYYYY = Null
    Err.Raise (Err.Number)
    Resume IsDateDDMMYYYYEXIT
End Function

Public Function IsCharIn(psInput As String, _
                            pcolAllowed As Collection)

On Error GoTo IsCharInERR

'We will allow two special words in the collection
'   'Alphabetic'
'   'Numeric'

Dim psAllowed As Variant
Dim ch As String * 1
Dim res As Integer

    ch = Left(psInput, 1)
    
    For Each psAllowed In pcolAllowed
        If UCase(psAllowed) = UCase("Alphabetic") Then
            res = res + ((Asc(UCase(ch)) >= 65 And Asc(UCase(ch)) <= 90))
        ElseIf UCase(psAllowed) = UCase("Numeric") Then
            res = res + (IsNumeric(ch))
        Else
            res = res + (UCase(ch) = UCase(Left(psAllowed, 1)))
        End If
    Next
    
    IsCharIn = (res <> 0)
    
IsCharInEXIT:
    Exit Function
IsCharInERR:
    IsCharIn = False
    Err.Raise Err.Number
    Resume IsCharInEXIT
End Function
