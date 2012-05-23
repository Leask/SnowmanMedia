VERSION 5.00
Begin VB.Form frmImage 
   Caption         =   "Editor  - Snowman Media Pictures Browser  1.0"
   ClientHeight    =   5040
   ClientLeft      =   5925
   ClientTop       =   3015
   ClientWidth     =   5400
   Icon            =   "看图.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   5400
   Begin VB.PictureBox picScrollBox 
      BackColor       =   &H00000000&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   0
      Width           =   5115
      Begin VB.HScrollBar hscScroll 
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   4185
         Width           =   4830
      End
      Begin VB.VScrollBar vscScroll 
         Height          =   4110
         Left            =   4860
         TabIndex        =   2
         Top             =   0
         Width           =   195
      End
      Begin VB.PictureBox picImage 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3345
         Left            =   0
         ScaleHeight     =   223
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   316
         TabIndex        =   1
         Top             =   0
         Width           =   4740
      End
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_sFIleName As String
Public m_sFIleTitle As String
Private m_bDirty As Boolean

Private WithEvents m_cImage As cImageProcessDIB
Attribute m_cImage.VB_VarHelpID = -1
Private m_cDib As New cDIBSection
Private m_cDibBuffer As New cDIBSection

Public Property Get ImageDibHDC() As Long
   ImageDibHDC = m_cDib.hdc
End Property

Public Sub Combine(ByRef fC As frmCombination)
   m_cDib.Create fC.NewImageWidth, fC.NewImageHeight
   m_cDibBuffer.Create m_cDib.Width, m_cDib.Height
   ' Copy the 1st source image to m_cDIb
   m_cDib.LoadPictureBlt Forms(fC.ImageSource(1)).ImageDibHDC
   ' Copy the 2nd to m_cDibBuffer
   m_cDibBuffer.LoadPictureBlt Forms(fC.ImageSource(2)).ImageDibHDC
   
   Select Case fC.CombinationType
   Case eAdd
      ' Add the images together:
      m_cImage.AddImages m_cDibBuffer, m_cDib, fC.Multiplier(2), fC.Offset(2), fC.Offset(2), fC.Offset(2), fC.Multiplier(1), fC.Offset(1), fC.Offset(1), fC.Offset(1)
   Case eDarkest
      m_cImage.AddDarkest m_cDibBuffer, m_cDib
   Case eLightest
      m_cImage.AddLightest m_cDibBuffer, m_cDib
   End Select
   FileName = "Image " & mfrmMain.NewImageIndex
   Me.Caption = "图片:" & FileName
   m_bDirty = True
   Render
End Sub

Public Property Get ImageWidth() As Long
   ImageWidth = m_cDib.Width
End Property
Public Property Get ImageHeight() As Long
   ImageHeight = m_cDib.Height
End Property


Public Sub ApplyPalette(ByVal sPalFile As String)
Dim cPal As New cPalette
   cPal.LoadFromFile sPalFile
   m_cImage.ApplyPalette m_cDib, m_cDibBuffer, cPal
   m_bDirty = True
   Render
End Sub

Public Sub Colourise(ByVal fHue As Single)
   ' Colourise takes hue (-1 to 5)
   m_cImage.Colourise m_cDib, fHue, 0.5
   m_bDirty = True
   Render
End Sub

Public Sub Lighten()
   ' Lighten takes percentage:
   m_cImage.Lighten m_cDib, 20
   m_bDirty = True
   Render
End Sub

Public Sub Fade()
   ' Fade 255 = no fading, 0 = all black
   m_cImage.Fade m_cDib, 240
   m_bDirty = True
   Render
End Sub

Public Sub BlackAndWhite()
    m_cImage.BlackAndWhite m_cDib, m_cDibBuffer
    m_bDirty = True
    Render
End Sub

Public Sub GrayScale()
    m_cImage.GrayScale m_cDib
    m_bDirty = True
    Render
End Sub

Public Sub NegativeImage()
    m_cImage.AddImages m_cDib, m_cDibBuffer, -1, -255, -255, -255, 0, 0, 0, 0
    m_cDibBuffer.PaintPicture m_cDib.hdc
    m_bDirty = True
    Render
End Sub

Public Sub AddNoise(ByVal bRandom As Boolean, ByVal lAmount As Long)
    m_cImage.AddNoise m_cDib, lAmount, bRandom
    m_bDirty = True
    Render
    m_bDirty = True
End Sub

Public Sub Resample(ByVal lW As Long, ByVal lH As Long)
Dim cDib As New cDIBSection
    If (lW <> m_cDib.Width) Or (lH <> m_cDib.Height) Then
        Set cDib = m_cDib.Resample(lH, lW)
        Set m_cDib = cDib
        m_cDibBuffer.Create m_cDib.Width, m_cDib.Height
        Render
    End If
    m_bDirty = True
End Sub

Public Sub Render()
    picImage.Width = m_cDib.Width * Screen.TwipsPerPixelX
    picImage.Height = m_cDib.Height * Screen.TwipsPerPixelY
    m_cDib.PaintPicture picImage.hdc
    picImage.Refresh
End Sub

Public Sub CopyImage()
    m_cDib.CopyToClipboard False
End Sub
Public Sub ProcessImage(ByVal eType As EFilterTypes)
    With m_cImage
        .FilterType = eType
        .ProcessImage m_cDib, m_cDibBuffer
        Render
        m_bDirty = True
    End With
End Sub
Public Sub LoadCustomFilter(ByRef cI As cImageProcessDIB)
Dim i As Long, j As Long
    With m_cImage
        .FilterType = eCustom
        .FilterWeight = cI.FilterWeight
        .FilterArraySize = cI.FilterArraySize
        For i = -cI.FilterArraySize \ 2 To cI.FilterArraySize \ 2
            For j = -cI.FilterArraySize \ 2 To cI.FilterArraySize \ 2
                .FilterValue(i, j) = cI.FilterValue(i, j)
            Next j
        Next i
    End With
End Sub

Public Property Get Dirty() As Boolean
    Dirty = m_bDirty
End Property
Public Function QuerySave() As Boolean
Dim eR As VbMsgBoxResult
    eR = MsgBox("图片:" & m_sFIleTitle & "已经被修改." & vbCrLf & vbCrLf & "你要保存对它所作的修改吗?", vbYesNoCancel Or vbQuestion)
    Select Case eR
    Case vbYes
        If (SaveFile()) Then
            QuerySave = True
        End If
    Case vbNo
        QuerySave = True
    Case vbCancel
        ' cancel..
    End Select
End Function

Public Function OpenFile(ByVal sFIle As String, Optional ByVal bIsTemp As Boolean = False) As Boolean
Dim sPic As StdPicture
On Error GoTo OpenFileError
    
    mfrmMain.SetStatus "正在加载图片:" & sFIle & "..."
    Set sPic = LoadPicture(sFIle)
    m_cDib.CreateFromPicture sPic
    m_cDibBuffer.Create m_cDib.Width, m_cDib.Height
    Render
    If (bIsTemp) Then
       sFIle = "Image " & mfrmMain.NewImageIndex
    End If
    Caption = "SmM. P.B.  1.0:" & sFIle
    FileName = sFIle
    If Not (bIsTemp) Then
        mfrmMain.SetStatus "已经打开: " & sFIle & ".", FileTitle, picImage.Width \ Screen.TwipsPerPixelX & " x " & picImage.Height \ Screen.TwipsPerPixelY
        mfrmMain.AddMRUFile sFIle
    End If
    picImage.Refresh
    picScrollBox_Resize
    OpenFile = True
    Exit Function
OpenFileError:
    MsgBox "Snowman Media Pictures Browser  1.0 在加载图片时发生错误:" & Err.Description, vbExclamation
    Exit Function
End Function
Public Function SaveFile() As Boolean
Dim sName As String
Dim iPos As Long
Dim i As Long
Dim c As New GCommonDialog

On Error GoTo SaveFileError

    ' Strip extenstion:
    For i = Len(m_sFIleName) To 1 Step -1
        If (Mid$(m_sFIleName, i, 1) = ".") Then
            iPos = i - 1
            Exit For
        End If
    Next i
    If (iPos > 1) Then
        sName = Left$(m_sFIleName, iPos) & ".bmp"
    Else
        sName = m_sFIleName & ".bmp"
    End If
    
    ' Ask to save:
    If c.VBGetSaveFileName(sName, , , "位图文件(*.BMP)|*.BMP|所有文件(*.*)|*.*", , , , "BMP", Me.hWnd) Then
        SavePicture picImage.Image, sName
        FileName = sName
        mfrmMain.AddMRUFile sName
        Caption = "图片:" & sName
        m_bDirty = False
    End If
    Exit Function

SaveFileError:
    MsgBox "Snowman Media Pictures Browser  1.0 在保存图片时发生错误." & Err.Description, vbExclamation
    Exit Function

End Function
Public Property Let FileName(ByVal sName As String)
Dim i As Long, iPos As Long
    m_sFIleName = sName
    For i = Len(sName) To 1 Step -1
        If Mid$(sName, i, 1) = "\" Then
            iPos = i + 1
            Exit For
        End If
    Next i
    If (iPos > 0) Then
        m_sFIleTitle = Mid$(sName, iPos)
    Else
        m_sFIleTitle = sName
    End If
    
End Property

Public Property Get FileName() As String
    FileName = m_sFIleName
End Property
Public Property Get FileTitle() As String
    FileTitle = m_sFIleTitle
End Property


Private Sub Form_Activate()
    mfrmMain.SetStatus , Me.FileTitle, picImage.Width \ Screen.TwipsPerPixelX & " x " & picImage.Height \ Screen.TwipsPerPixelY
End Sub

Private Sub Form_Load()
    '
    Set m_cImage = New cImageProcessDIB
End Sub


Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
      On Error Resume Next
      picScrollBox.Move 2 * Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY, Me.ScaleWidth - 4 * Screen.TwipsPerPixelX, Me.ScaleHeight - 4 * Screen.TwipsPerPixelY
    End If
End Sub

Private Sub hscScroll_Change()
    picImage.Left = -Screen.TwipsPerPixelY * hscScroll.Value
End Sub

Private Sub hscScroll_Scroll()
    hscScroll_Change
End Sub

Private Sub m_cImage_Complete(ByVal lTimeMs As Long)
    mfrmMain.ShowProgress = False
    mfrmMain.SetStatus "Complete.  Time = " & lTimeMs
End Sub

Private Sub m_cImage_InitProgress(ByVal lMax As Long)
    mfrmMain.ProgressMax = lMax
    mfrmMain.ProgressValue = 0
    mfrmMain.ShowProgress = True
End Sub

Private Sub m_cImage_Progress(ByVal lPosition As Long)
    mfrmMain.ProgressValue = lPosition
End Sub

Private Sub picScrollBox_Resize()
    On Error Resume Next
    hscScroll.Visible = (picScrollBox.ScaleWidth - vscScroll.Width < picImage.Width)
    vscScroll.Visible = (picScrollBox.ScaleHeight - hscScroll.Height < picImage.Height)
    If (hscScroll.Visible) Then
        hscScroll.Max = (picImage.Width - picScrollBox.ScaleWidth + vscScroll.Width * Abs(vscScroll.Visible)) \ Screen.TwipsPerPixelX
        hscScroll.SmallChange = 32
        hscScroll.Move 0, picScrollBox.ScaleHeight - hscScroll.Height, picScrollBox.ScaleWidth - (vscScroll.Width * Abs(vscScroll.Visible))
    End If
    If (vscScroll.Visible) Then
        vscScroll.Max = (picImage.Height - picScrollBox.ScaleHeight + hscScroll.Height * Abs(hscScroll.Visible)) \ Screen.TwipsPerPixelY
        vscScroll.SmallChange = 32
        vscScroll.Move picScrollBox.ScaleWidth - vscScroll.Width, 0, vscScroll.Width, picScrollBox.ScaleHeight - (hscScroll.Height * Abs(hscScroll.Visible))
    End If
End Sub

Private Sub vscScroll_Change()
    picImage.Top = -Screen.TwipsPerPixelY * vscScroll.Value
End Sub

Private Sub vscScroll_Scroll()
    vscScroll_Change
End Sub
