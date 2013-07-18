VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PatternFade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const SRCCOPY& = &HCC0020
Private Const SRCAND& = &H8800C6
Dim BitBrush(20) As BitBrush
Dim m_pict1 As Control
Dim m_pict2 As Control
Dim m_pict3 As Control
Public Property Set pic1(Acontrol As Control)
    Set m_pict1 = Acontrol
End Property
Public Property Set pic2(Acontrol As Control)
    Set m_pict2 = Acontrol
End Property
Public Property Set pic3(Acontrol As Control)
    Set m_pict3 = Acontrol
End Property
Private Sub CreatePatterns()
    Dim Cnt As Long
    Dim N As Long
    Dim Pat As String
    For Cnt = 1 To 20
        For N = 1 To 8
                Select Case Cnt
                   Case 1
                      If N = 1 Then Pat$ = "00000000"
                      If N = 2 Then Pat$ = "00000000"
                      If N = 3 Then Pat$ = "00000000"
                      If N = 4 Then Pat$ = "00000000"
                      If N = 5 Then Pat$ = "00000000"
                      If N = 6 Then Pat$ = "00000000"
                      If N = 7 Then Pat$ = "00000000"
                      If N = 8 Then Pat$ = "00000000"
                  Case 2
                      If N = 1 Then Pat$ = "10001000"
                      If N = 2 Then Pat$ = "00000000"
                      If N = 3 Then Pat$ = "00000000"
                      If N = 4 Then Pat$ = "00000000"
                      If N = 5 Then Pat$ = "10001000"
                      If N = 6 Then Pat$ = "00000000"
                      If N = 7 Then Pat$ = "00000000"
                      If N = 8 Then Pat$ = "00000000"
                  Case 3
                      If N = 1 Then Pat$ = "10001000"
                      If N = 2 Then Pat$ = "00000000"
                      If N = 3 Then Pat$ = "00100010"
                      If N = 4 Then Pat$ = "00000000"
                      If N = 5 Then Pat$ = "10001000"
                      If N = 6 Then Pat$ = "00000000"
                      If N = 7 Then Pat$ = "00100010"
                      If N = 8 Then Pat$ = "00000000"
                  Case 4
                      If N = 1 Then Pat$ = "10001000"
                      If N = 2 Then Pat$ = "00000000"
                      If N = 3 Then Pat$ = "10101010"
                      If N = 4 Then Pat$ = "00000000"
                      If N = 5 Then Pat$ = "10001000"
                      If N = 6 Then Pat$ = "00000000"
                      If N = 7 Then Pat$ = "10101010"
                      If N = 8 Then Pat$ = "00000000"
                  Case 5
                      If N = 1 Then Pat$ = "10101010"
                      If N = 2 Then Pat$ = "00000000"
                      If N = 3 Then Pat$ = "10101010"
                      If N = 4 Then Pat$ = "00000000"
                      If N = 5 Then Pat$ = "10101010"
                      If N = 6 Then Pat$ = "00000000"
                      If N = 7 Then Pat$ = "10101010"
                      If N = 8 Then Pat$ = "00000000"
                  Case 6
                      If N = 1 Then Pat$ = "10101010"
                      If N = 2 Then Pat$ = "00010001"
                      If N = 3 Then Pat$ = "10101010"
                      If N = 4 Then Pat$ = "00000000"
                      If N = 5 Then Pat$ = "10101010"
                      If N = 6 Then Pat$ = "00010001"
                      If N = 7 Then Pat$ = "10101010"
                      If N = 8 Then Pat$ = "00000000"
                  Case 7
                      If N = 1 Then Pat$ = "10101010"
                      If N = 2 Then Pat$ = "00010001"
                      If N = 3 Then Pat$ = "10101010"
                      If N = 4 Then Pat$ = "01000100"
                      If N = 5 Then Pat$ = "10101010"
                      If N = 6 Then Pat$ = "00010001"
                      If N = 7 Then Pat$ = "10101010"
                      If N = 8 Then Pat$ = "01000100"
                  Case 8
                      If N = 1 Then Pat$ = "10101010"
                      If N = 2 Then Pat$ = "01010101"
                      If N = 3 Then Pat$ = "10101010"
                      If N = 4 Then Pat$ = "01000100"
                      If N = 5 Then Pat$ = "10101010"
                      If N = 6 Then Pat$ = "01010101"
                      If N = 7 Then Pat$ = "10101010"
                      If N = 8 Then Pat$ = "01000100"
                  Case 9
                      If N = 1 Then Pat$ = "10101010"
                      If N = 2 Then Pat$ = "01010101"
                      If N = 3 Then Pat$ = "10101010"
                      If N = 4 Then Pat$ = "01010101"
                      If N = 5 Then Pat$ = "10101010"
                      If N = 6 Then Pat$ = "01010101"
                      If N = 7 Then Pat$ = "10101010"
                      If N = 8 Then Pat$ = "01010101"
                  Case 10
                      If N = 1 Then Pat$ = "11101110"
                      If N = 2 Then Pat$ = "01010101"
                      If N = 3 Then Pat$ = "10101010"
                      If N = 4 Then Pat$ = "01010101"
                      If N = 5 Then Pat$ = "11101110"
                      If N = 6 Then Pat$ = "01010101"
                      If N = 7 Then Pat$ = "10101010"
                      If N = 8 Then Pat$ = "01010101"
                  Case 11
                      If N = 1 Then Pat$ = "11101110"
                      If N = 2 Then Pat$ = "01010101"
                      If N = 3 Then Pat$ = "10111011"
                      If N = 4 Then Pat$ = "01010101"
                      If N = 5 Then Pat$ = "11101110"
                      If N = 6 Then Pat$ = "01010101"
                      If N = 7 Then Pat$ = "10111011"
                      If N = 8 Then Pat$ = "01010101"
                  Case 12
                      If N = 1 Then Pat$ = "11101110"
                      If N = 2 Then Pat$ = "01010101"
                      If N = 3 Then Pat$ = "11111111"
                      If N = 4 Then Pat$ = "01010101"
                      If N = 5 Then Pat$ = "11101110"
                      If N = 6 Then Pat$ = "01010101"
                      If N = 7 Then Pat$ = "11111111"
                      If N = 8 Then Pat$ = "01010101"
                  Case 13
                      If N = 1 Then Pat$ = "11101110"
                      If N = 2 Then Pat$ = "01110111"
                      If N = 3 Then Pat$ = "11111111"
                      If N = 4 Then Pat$ = "01010101"
                      If N = 5 Then Pat$ = "11101110"
                      If N = 6 Then Pat$ = "01110111"
                      If N = 7 Then Pat$ = "11111111"
                      If N = 8 Then Pat$ = "01010101"
                  Case 14
                      If N = 1 Then Pat$ = "11101110"
                      If N = 2 Then Pat$ = "01110111"
                      If N = 3 Then Pat$ = "11111111"
                      If N = 4 Then Pat$ = "11011101"
                      If N = 5 Then Pat$ = "11101110"
                      If N = 6 Then Pat$ = "01110111"
                      If N = 7 Then Pat$ = "11111111"
                      If N = 8 Then Pat$ = "11011101"
                  Case 15
                      If N = 1 Then Pat$ = "11111111"
                      If N = 2 Then Pat$ = "01110111"
                      If N = 3 Then Pat$ = "11111111"
                      If N = 4 Then Pat$ = "11011101"
                      If N = 5 Then Pat$ = "11111111"
                      If N = 6 Then Pat$ = "01110111"
                      If N = 7 Then Pat$ = "11111111"
                      If N = 8 Then Pat$ = "11011101"
                   Case 16
                      If N = 1 Then Pat$ = "11111111"
                      If N = 2 Then Pat$ = "11111111"
                      If N = 3 Then Pat$ = "11111111"
                      If N = 4 Then Pat$ = "11011101"
                      If N = 5 Then Pat$ = "11111111"
                      If N = 6 Then Pat$ = "11111111"
                      If N = 7 Then Pat$ = "11111111"
                      If N = 8 Then Pat$ = "11011101"
                    Case 17
                      If N = 1 Then Pat$ = "11111111"
                      If N = 2 Then Pat$ = "11111111"
                      If N = 3 Then Pat$ = "11111111"
                      If N = 4 Then Pat$ = "11111101"
                      If N = 5 Then Pat$ = "11111111"
                      If N = 6 Then Pat$ = "11111111"
                      If N = 7 Then Pat$ = "11111111"
                      If N = 8 Then Pat$ = "11011101"
                   Case 18
                      If N = 1 Then Pat$ = "11111111"
                      If N = 2 Then Pat$ = "11111111"
                      If N = 3 Then Pat$ = "11111111"
                      If N = 4 Then Pat$ = "11111101"
                      If N = 5 Then Pat$ = "11111111"
                      If N = 6 Then Pat$ = "11111111"
                      If N = 7 Then Pat$ = "11111111"
                      If N = 8 Then Pat$ = "11011111"
                   Case 19
                      If N = 1 Then Pat$ = "11111111"
                      If N = 2 Then Pat$ = "11111111"
                      If N = 3 Then Pat$ = "11111111"
                      If N = 4 Then Pat$ = "11111101"
                      If N = 5 Then Pat$ = "11111111"
                      If N = 6 Then Pat$ = "11111111"
                      If N = 7 Then Pat$ = "11111111"
                      If N = 8 Then Pat$ = "11111111"
                  Case 20
                      If N = 1 Then Pat$ = "11111111"
                      If N = 2 Then Pat$ = "11111111"
                      If N = 3 Then Pat$ = "11111111"
                      If N = 4 Then Pat$ = "11111111"
                      If N = 5 Then Pat$ = "11111111"
                      If N = 6 Then Pat$ = "11111111"
                      If N = 7 Then Pat$ = "11111111"
                      If N = 8 Then Pat$ = "11111111"
                 End Select
                 BitBrush(Cnt).SetPattern Pat$, N
        Next N
        BitBrush(Cnt).SetuBitmap 255, 255, 255, 0, 0, 0
        BitBrush(Cnt).BuildBitmap
    Next Cnt
End Sub
Public Sub FadeIn(Time As Integer)
    Dim Cnt As Integer
    Dim throw As Long
    CreatePatterns
    For Cnt = 20 To 1 Step -1
        BitBrush(Cnt).ShowPattern
        throw = BitBlt(m_pict2.hDC, 0, 0, m_pict1.ScaleWidth, m_pict1.ScaleHeight, m_pict1.hDC, 0, 0, SRCAND)
        throw = BitBlt(m_pict3.hDC, 0, 0, m_pict3.ScaleWidth, m_pict3.ScaleHeight, m_pict2.hDC, 0, 0, SRCCOPY)
        Sleep Time
        DoEvents
    Next Cnt
End Sub
Public Sub FadeOut(Time As Integer)
    Dim Cnt As Integer
    Dim throw As Integer
    CreatePatterns
     For Cnt = 1 To 20
        BitBrush(Cnt).ShowPattern
        throw = BitBlt(m_pict2.hDC, 0, 0, m_pict1.ScaleWidth, m_pict1.ScaleHeight, m_pict1.hDC, 0, 0, SRCAND)
        throw = BitBlt(m_pict3.hDC, 0, 0, m_pict3.ScaleWidth, m_pict3.ScaleHeight, m_pict2.hDC, 0, 0, SRCCOPY)
        Sleep Time
        DoEvents
    Next Cnt
End Sub
Public Sub Setup()
    Dim i As Integer
    m_pict1.Width = m_pict3.Width
    m_pict1.Height = m_pict3.Height
    m_pict2.Width = m_pict3.Width
    m_pict2.Height = m_pict3.Height
    m_pict1.AutoRedraw = True
    m_pict2.AutoRedraw = True
    For i = 1 To 20
         Set BitBrush(i) = New BitBrush
         Set BitBrush(i).DispPict = m_pict2
    Next
End Sub
