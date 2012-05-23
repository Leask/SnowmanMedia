VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(1) As RGBQUAD
End Type
Private Declare Function SelectObject& Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long)
Private Declare Function CreateDIBitmap& Lib "gdi32" (ByVal hDC As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO, ByVal wUsage As Long)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Private Declare Function CreatePatternBrush& Lib "gdi32" (ByVal hBitmap As Long)
Private Declare Function PatBlt& Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long)
Private Const DIB_RGB_COLORS& = 0
Private Const CBM_INIT& = &H4
Private Const PATCOPY& = &HF00021
Private Const BI_RGB& = 0&
Dim m_BitInfoH As BITMAPINFOHEADER
Dim m_BitInfo As BITMAPINFO
Dim da(32) As Byte
Dim m_Hbr As Long
Dim m_OldBP As Long
Private m_DispPict As Object
Private m_Array(8) As String * 8
Public Property Set DispPict(Acontrol As Object)
     Set m_DispPict = Acontrol
End Property
Public Sub SetuBitmap(r1, g1, b1, r2, g2, b2)
    m_BitInfoH.biSize = 40
    m_BitInfoH.biWidth = 8
    m_BitInfoH.biHeight = 8
    m_BitInfoH.biPlanes = 1
    m_BitInfoH.biBitCount = 1
    m_BitInfoH.biCompression = BI_RGB
    m_BitInfoH.biSizeImage = 0
    m_BitInfoH.biXPelsPerMeter = 0
    m_BitInfoH.biYPelsPerMeter = 0
    m_BitInfoH.biClrUsed = 2
    m_BitInfoH.biClrImportant = 0
    m_BitInfo.bmiColors(0).rgbBlue = r1
    m_BitInfo.bmiColors(0).rgbGreen = g1
    m_BitInfo.bmiColors(0).rgbRed = b1
    m_BitInfo.bmiColors(0).rgbReserved = 1
    m_BitInfo.bmiColors(1).rgbBlue = r2
    m_BitInfo.bmiColors(1).rgbGreen = g2
    m_BitInfo.bmiColors(1).rgbRed = b2
    m_BitInfo.bmiColors(1).rgbReserved = 0
 End Sub
Public Sub BuildBitmap()
    Dim Counter As Integer, V As Integer, c As Integer
    Dim CompBitmap As Long
    Dim dl As Long
    For Counter = 1 To 8
        V = 0
        For c = 0 To 7
            If Mid$(m_Array(Counter), c + 1, 1) = "1" Then V = V + 2 ^ c
        Next c
        da(Counter * 4 - 4) = CByte(V)
    Next Counter
    m_BitInfo.bmiHeader = m_BitInfoH
    CompBitmap = CreateDIBitmap(m_DispPict.hDC, m_BitInfoH, CBM_INIT, da(0), _
                                        m_BitInfo, DIB_RGB_COLORS)
    m_Hbr = CreatePatternBrush(CompBitmap)
    dl& = DeleteObject(CompBitmap)
End Sub
Public Sub DeleteBrush()
    Dim throw As Long
    throw& = SelectObject(m_DispPict.hDC, m_OldBP)
    throw& = DeleteObject(m_Hbr)
End Sub
Public Sub SelectBrush()
   m_OldBP = SelectObject(m_DispPict.hDC, m_Hbr)
End Sub
Public Sub ShowPattern()
    Dim throw As Long
    m_OldBP = SelectObject(m_DispPict.hDC, m_Hbr)
    throw& = PatBlt(m_DispPict.hDC, 0, 0, m_DispPict.ScaleWidth, m_DispPict.ScaleHeight, PATCOPY)
    throw& = SelectObject(m_DispPict.hDC, m_OldBP)
    throw& = DeleteObject(m_Hbr)
End Sub
Public Sub SetPattern(s, Index)
    m_Array(Index) = s
End Sub
Private Sub Class_Initialize()
    If m_Hbr Then DeleteBrush
End Sub
