VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmThumbs 
   Caption         =   "���  �� Snowman Media Pictures Browser  1.0"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   Icon            =   "���.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   286
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   2.45745e5
      TabIndex        =   13
      Top             =   0
      Width           =   2.45745e5
      Begin VB.PictureBox picCount 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7290
         ScaleHeight     =   255
         ScaleWidth      =   2775
         TabIndex        =   24
         Top             =   45
         Width           =   2805
         Begin VB.Label lblCount 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            Caption         =   "���ļ��й���:ͼƬxx��."
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   45
            TabIndex        =   25
            Top             =   45
            Width           =   1980
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5760
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   21
         Top             =   45
         Width           =   1365
         Begin VB.Label Label6 
            BackColor       =   &H0000FFFF&
            Caption         =   "�˳�(&X)"
            ForeColor       =   &H00FF0000&
            Height          =   690
            Left            =   45
            TabIndex        =   22
            Top             =   45
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4320
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   19
         Top             =   45
         Width           =   1365
         Begin VB.Label Label5 
            BackColor       =   &H0000FFFF&
            Caption         =   "�õ�Ƭ��ʽ(&M)"
            ForeColor       =   &H00FF0000&
            Height          =   780
            Left            =   45
            TabIndex        =   20
            Top             =   45
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2880
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   17
         Top             =   45
         Width           =   1365
         Begin VB.Label Label4 
            BackColor       =   &H0000FFFF&
            Caption         =   "ˢ��(&R)"
            ForeColor       =   &H00FF0000&
            Height          =   645
            Left            =   45
            TabIndex        =   18
            Top             =   45
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   16
         Top             =   45
         Width           =   1365
         Begin VB.Label Label3 
            BackColor       =   &H0000FFFF&
            Caption         =   "�༭(&E)"
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   45
            TabIndex        =   23
            Top             =   45
            Width           =   2130
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   1335
         TabIndex        =   14
         Top             =   45
         Width           =   1365
         Begin VB.Label Label2 
            BackColor       =   &H0000FFFF&
            Caption         =   "���(&B)"
            ForeColor       =   &H00FF0000&
            Height          =   510
            Left            =   45
            TabIndex        =   15
            Top             =   45
            Width           =   1455
         End
      End
   End
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      Height          =   285
      Left            =   3645
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   8
      Top             =   8730
      Width           =   915
      Begin VB.PictureBox picProgressSlide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   0
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   9
         Top             =   0
         Width           =   255
      End
   End
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   3960
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   529
            MinWidth        =   529
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14605
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picThumb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   2565
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   4
      Top             =   6975
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.FileListBox filHidden 
      Height          =   450
      Left            =   3645
      Pattern         =   "*.bmp;*.dib;*.rle;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur"
      TabIndex        =   6
      Top             =   7110
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox picLoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1230
      Left            =   2565
      ScaleHeight     =   82
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   82
      TabIndex        =   3
      Top             =   5760
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.PictureBox picFrame 
      Height          =   2175
      Left            =   450
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   132
      TabIndex        =   0
      Top             =   5760
      Width           =   2040
      Begin VB.PictureBox picSlide 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   90
         ScaleHeight     =   115
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   106
         TabIndex        =   2
         Top             =   90
         Width           =   1590
         Begin VB.OptionButton optThumb 
            Height          =   1455
            Index           =   0
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   45
            Width           =   1230
         End
      End
      Begin VB.VScrollBar vsbSlide 
         Height          =   2130
         Left            =   1755
         TabIndex        =   1
         Top             =   0
         Width           =   240
      End
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   635
      ButtonWidth     =   2937
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "���(&B)"
            Key             =   "Browse"
            Object.ToolTipText     =   "Select a Folder"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�༭(&E)"
            Key             =   "Edit"
            Object.ToolTipText     =   "Open selected image in an Editor"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ˢ��(&R)"
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh Thumbnails"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�õ�Ƭ��ʽ(&M)"
            Key             =   "power"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "�˳�(&X)"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit Program"
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5265
         ScaleHeight     =   285
         ScaleWidth      =   1815
         TabIndex        =   11
         Top             =   405
         Visible         =   0   'False
         Width           =   1815
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "lblCount"
            Height          =   195
            Left            =   0
            TabIndex        =   12
            Top             =   135
            Visible         =   0   'False
            Width           =   570
         End
      End
   End
End
Attribute VB_Name = "frmThumbs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************
'*******************************************************************
'**SuperViewer1.0                                                 **
'**Developed by:  Mallesh Bonigala                                **
'**Credits:   Some parts of the program contain source borrowed   **
'**           from vbcode.com and planet-source-code.com          **
'**My sincere thanks to all the guys whose techniques I have in-  **
'**corporated in this program.                                    **
'*******************************************************************
'*******************************************************************
Option Explicit

Private Type PointAPI
    X  As Long
    Y  As Long
End Type

Private mbActive                As Boolean
Private mlCurThumb              As Long
Private Const SRCCOPY           As Long = &HCC0020
Private Const STRETCH_HALFTONE  As Long = &H4&
Private Const SW_RESTORE        As Long = &H9&

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lpPt As PointAPI) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function UnrealizeObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Browse(Optional ByVal bDontShowBrowser As Boolean)

Dim lRet    As Long
Dim sPath   As String

    sPath = GetInitEntryB("Main", "Last Path", "C:\")
   ' If Not bDontShowBrowser Then
        sPath = BrowseForFolder(Me.hWnd, "��ѡ����Ҫ������ļ���, Snowman Media Pictures Browser  1.0 ��Ϊ�����.", sPath)
    'End If
    If Len(sPath) > 0 Then
        On Error Resume Next
        filHidden.Path = sPath
        If Err.Number = 0 Then
            CreateThumbs
            lRet = SetInitEntry("Main", "Last Path", sPath)
        End If
        
    End If
Me.Caption = sPath + "\  - Snowman Media Pictures Browser  1.0"
End Sub

Private Sub CreateThumbPic(picSource As PictureBox, picThumb As PictureBox)

'This sub uses the halftone stretch mode, which produces the highest
'quality possible, when stretching the bitmap.

Dim lRet            As Long
Dim lLeft           As Long
Dim lTop            As Long
Dim lWidth          As Long
Dim lHeight         As Long
Dim lForeColor      As Long
Dim hBrush          As Long
Dim hDummyBrush     As Long
Dim lOrigMode       As Long
Dim fScale          As Single
Dim uBrushOrigPt    As PointAPI

    picThumb.Width = 64
    picThumb.Height = 64
    picThumb.BackColor = vbButtonFace
    picThumb.AutoRedraw = True
    picThumb.Cls
    
    If picSource.Width <= picThumb.Width - 2 And picSource.Height <= picThumb.Height - 2 Then
        fScale = 1
    Else
        fScale = IIf(picSource.Width > picSource.Height, (picThumb.Width - 2) / picSource.Width, (picThumb.Height - 2) / picSource.Height)
    End If
    lWidth = picSource.Width * fScale
    lHeight = picSource.Height * fScale
    lLeft = Int((picThumb.Width - lWidth) / 2)
    lTop = Int((picThumb.Height - lHeight) / 2)
    
    'Store the original ForeColor
    lForeColor = picThumb.ForeColor
    
    'Set picEdit's stretch mode to halftone (this may cause misalignment of the brush)
    lOrigMode = SetStretchBltMode(picThumb.hDC, STRETCH_HALFTONE)
    
    'Realign the brush...
    'Get picEdit's brush by selecting a dummy brush into the DC
    hDummyBrush = CreateSolidBrush(lForeColor)
    hBrush = SelectObject(picThumb.hDC, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    lRet = UnrealizeObject(hBrush)
    'Set picEdit's brush alignment coordinates to the left-top of the bitmap
    lRet = SetBrushOrgEx(picThumb.hDC, lLeft, lTop, uBrushOrigPt)
    'Now put the original brush back into the DC at the new alignment
    hDummyBrush = SelectObject(picThumb.hDC, hBrush)
    
    'Stretch the bitmap
    lRet = StretchBlt(picThumb.hDC, lLeft, lTop, lWidth, lHeight, _
            picSource.hDC, 0, 0, picSource.Width, picSource.Height, SRCCOPY)
    
    'Set the stretch mode back to it's original mode
    lRet = SetStretchBltMode(picThumb.hDC, lOrigMode)
    
    'Reset the original alignment of the brush...
    'Get picEdit's brush by selecting the dummy brush back into the DC
    hBrush = SelectObject(picThumb.hDC, hDummyBrush)
    'Reset the brush (This will force windows to realign it when it's put back)
    lRet = UnrealizeObject(hBrush)
    'Set the brush alignment back to the original coordinates
    lRet = SetBrushOrgEx(picThumb.hDC, uBrushOrigPt.X, uBrushOrigPt.Y, uBrushOrigPt)
    'Now put the original brush back into picEdit's DC at the original alignment
    hDummyBrush = SelectObject(picThumb.hDC, hBrush)
    'Get rid of the dummy brush
    lRet = DeleteObject(hDummyBrush)
    
    'Restore the original ForeColor
    picThumb.ForeColor = lForeColor

    picThumb.Line (lLeft - 1, lTop - 1)-Step(lWidth + 1, lHeight + 1), &H0&, B
    
End Sub

Private Sub CreateThumbs()

Dim iMaxLen As Integer
Dim X       As Long
Dim Y       As Long
Dim lIdx    As Long
Dim lPicCnt As Long
Dim lFilCnt As Long
Dim sPath   As String
Dim sText   As String

    Screen.MousePointer = vbHourglass
    filHidden.Refresh
    
    picSlide.Move 0, 0, optThumb(0).Width, optThumb(0).Height
    picSlide.Visible = False
    picSlide.BackColor = vbButtonFace
    Set picSlide.Font = optThumb(0).Font
    While optThumb.Count > 1
        Unload optThumb(optThumb.Count - 1)
    Wend
    DoEvents
    On Error Resume Next
    sPath = filHidden.Path
    sPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", "")
    lFilCnt = filHidden.ListCount
    lblCount.Caption = "���ļ��й���ͼƬ��:" & Format$(lFilCnt, "#,##0")
    If Len(sPath) > 0 Then
        Call StartProgress
        For lIdx = 0 To filHidden.ListCount - 1
            Call UpdateProgress((CSng(lIdx + 1) / CSng(lFilCnt)) * 100, filHidden.List(lIdx))
            Set picLoad.Picture = LoadPicture()
            picLoad.Cls
            Err.Clear
            If InStr(1, LCase$(filHidden.List(lIdx)), ".ico") > 0 _
              Or InStr(1, LCase$(filHidden.List(lIdx)), ".cur") > 0 Then
                Set picLoad.Picture = LoadPicture(sPath & filHidden.List(lIdx), vbLPLargeShell, vbLPDefault)
            Else
                Set picLoad.Picture = LoadPicture(sPath & filHidden.List(lIdx))
            End If
            If Err.Number = 0 Then
                Call CreateThumbPic(picLoad, picThumb)
                If lPicCnt > 0 Then
                    Load optThumb(lPicCnt)
                    Set optThumb(lPicCnt).Container = picSlide
                End If
                optThumb(lPicCnt).Tag = filHidden.List(lIdx)
                Set optThumb(lPicCnt).Picture = picThumb.Image
                sText = filHidden.List(lIdx)
                iMaxLen = optThumb(lPicCnt).Width - 15
                If picSlide.TextWidth(sText) > iMaxLen Then
                    iMaxLen = iMaxLen - picSlide.TextWidth("...")
                End If
                While picSlide.TextWidth(sText) > iMaxLen
                    sText = Left$(sText, Len(sText) - 1)
                Wend
                If iMaxLen < optThumb(lPicCnt).Width - 15 Then
                    sText = sText & "..."
                End If
                optThumb(lPicCnt).Caption = sText
                optThumb(lPicCnt).Visible = True
                lPicCnt = lPicCnt + 1
            End If
        Next lIdx
        
        picProgress.Visible = False
        
        'Free the unneeded resources
        Set picLoad.Picture = LoadPicture()
        Set picThumb.Picture = LoadPicture()
        optThumb(0).Value = True
        mlCurThumb = 0
        Call Form_Resize
        picSlide.Visible = True
        lblCount.Caption = "���ļ��й���ͼƬ��:" & Format$(lPicCnt, "#,##0")
        
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Edit()

Dim lRet    As Long
Dim sFIle   As String

    sFIle = filHidden.Path
    If Len(sFIle) > 0 Then
        sFIle = sFIle & IIf(Right$(sFIle, 1) <> "\", "\", "")
        sFIle = sFIle & optThumb(mlCurThumb).Tag
       ' lRet = ShellExecute(Me.hWnd, "Open", sFIle, &H0&, &H0&, SW_RESTORE)
    End If

Dim f As New frmImage
        If (f.OpenFile(sFIle)) Then
            f.Show
        Else
            Unload f
        End If
    









End Sub

Private Sub StartProgress()

    With picProgress
        .Cls
        .BackColor = vbButtonFace
        .ForeColor = vbButtonText
        .Move sbrMain.Left + sbrMain.Panels(2).Left, sbrMain.Top + 1, _
            sbrMain.Panels(2).Width, sbrMain.Height - 1
    End With
    
    With picProgressSlide
        .Cls
        .BackColor = vbHighlight
        .ForeColor = vbHighlightText
        .Move 0, 0, 1, picProgress.ScaleHeight
    End With
    
    picProgress.Visible = True
    
End Sub

Private Sub UpdateProgress(ByVal iPercent As Integer, ByVal sCaption As String)

Dim lTextTop    As Long

    picProgress.Cls
    picProgressSlide.Cls
    picProgressSlide.Width = picProgress.ScaleWidth * (CSng(iPercent) / 100!)
    lTextTop = (picProgress.ScaleHeight - picProgress.TextHeight(sCaption)) / 2
    picProgress.CurrentX = 3
    picProgress.CurrentY = lTextTop
    picProgress.Print sCaption
    picProgressSlide.CurrentX = 3
    picProgressSlide.CurrentY = lTextTop
    picProgressSlide.Print sCaption
    DoEvents
    
End Sub

Private Sub Command1_Click()
 
End Sub

Private Sub Command2_Click()
 
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub filHidden_PathChange()

Dim sPath As String

    sPath = filHidden.Path
    If Len(sPath) > 0 Then
        sPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", "")
    End If
    sbrMain.Panels(2).Text = sPath
    lblCount.Caption = "���ļ��в�����ͼƬ�ļ�."
    
End Sub


Private Sub Form_Activate()
  
 If Not mbActive Then
        picSlide.Visible = False
     'DoEvents
        mbActive = True
       Call Browse(True)
    End If
  
End Sub

Private Sub Form_Load()
Dim sPos    As String
Dim saPos() As String
Dim laPos() As Long

    'Get the window position
    'sPos = GetInitEntryB("Main", "Window Position", sPos)
    'If Len(sPos) > 0 Then
     '   saPos = Split(sPos, ", ")
    'End If
    
    'Just in case saPos() is Empty
    ReDim Preserve saPos(3)
    ReDim laPos(3)
    laPos(0) = IIf(Len(saPos(0)) = 0, Me.Left, Val(Trim$(saPos(0))))
    laPos(1) = IIf(Len(saPos(1)) = 0, Me.Top, Val(Trim$(saPos(1))))
    laPos(2) = IIf(Len(saPos(2)) = 0, Me.Width, Val(Trim$(saPos(2))))
    laPos(3) = IIf(Len(saPos(3)) = 0, Me.Height, Val(Trim$(saPos(3))))
    
    Me.Move laPos(0), laPos(1), laPos(2), laPos(3)
    If CBool(GetInitEntryB("Main", "Maximized", CStr(False))) Then
        Me.WindowState = vbMaximized
    End If
    
    
    
    
    
    
    
End Sub


Private Sub Form_Resize()
On Error Resume Next
Dim X       As Long
Dim Y       As Long
Dim lIdx    As Long
Dim lCols   As Long
        
    If Me.WindowState <> vbMinimized Then
        If Me.Width < 346 * Screen.TwipsPerPixelX Then
        Me.Width = 346 * Screen.TwipsPerPixelX
        ElseIf Me.Height < 378 * Screen.TwipsPerPixelY Then
     Me.Height = 378 * Screen.TwipsPerPixelY
        Else
            picFrame.Move 0, tbrMain.Height, Me.ScaleWidth, Me.ScaleHeight - tbrMain.Height - sbrMain.Height
            vsbSlide.Move picFrame.ScaleWidth - vsbSlide.Width, 0, vsbSlide.Width, picFrame.ScaleHeight
            lCols = Int((picFrame.ScaleWidth - vsbSlide.Width) / optThumb(0).Width)
            For lIdx = 0 To optThumb.Count - 1
                X = (lIdx Mod lCols) * optThumb(0).Width
                Y = Int(lIdx / lCols) * optThumb(0).Height
                optThumb(lIdx).Move X, Y
            Next lIdx
            picSlide.Width = lCols * optThumb(0).Width
            picSlide.Height = Int(optThumb.Count / lCols) * optThumb(0).Height
            If Int(optThumb.Count / lCols) < (optThumb.Count / lCols) Then
                picSlide.Height = picSlide.Height + optThumb(0).Height
            End If
            vsbSlide.Value = 0
            vsbSlide.Max = picSlide.Height - picFrame.ScaleHeight
            If vsbSlide.Max < 0 Then
                vsbSlide.Max = 0
                vsbSlide.Enabled = False
            Else
                vsbSlide.Enabled = True
                vsbSlide.SmallChange = optThumb(0).Height
                vsbSlide.LargeChange = picFrame.ScaleHeight
            End If
        End If
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

Dim lIdx As Long
Dim lRet As Long
Dim sPos As String

    For lIdx = 1 To optThumb.Count - 1
        Unload optThumb(lIdx)
    Next lIdx
    
    If Me.WindowState = vbNormal Then
        sPos = CStr(Me.Left) & ", " & CStr(Me.Top) & ", " & CStr(Me.Width) & ", " & CStr(Me.Height)
        lRet = SetInitEntry("Main", "Window Position", sPos)
    End If
    lRet = SetInitEntry("Main", "Maximized", CStr(Me.WindowState = vbMaximized))
    
End Sub



Private Sub Label2_Click()
Call Browse
End Sub

Private Sub Label3_Click()
 Call Edit
End Sub

Private Sub Label4_Click()
Call CreateThumbs
End Sub

Private Sub Label5_Click()
Form1.Dir1.Path = filHidden.Path
Form1.Show
End Sub

Private Sub Label6_Click()
Unload Me
End Sub

Private Sub optThumb_Click(Index As Integer)

Dim sPath As String

    sPath = filHidden.Path
    If Len(sPath) > 0 Then
        sPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", "")
    End If
    sbrMain.Panels(2).Text = sPath & optThumb(Index).Tag
    mlCurThumb = Index

End Sub

Private Sub optThumb_DblClick(Index As Integer)

    Call Edit
    
End Sub




Private Sub picFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture3.BackColor = &HFFFF&
Picture4.BackColor = &HFFFF&
Picture5.BackColor = &HFFFF&
Picture6.BackColor = &HFFFF&
Picture7.BackColor = &HFFFF&
Label2.ForeColor = &HFF0000
Label2.BackColor = &HFFFF&
Label3.ForeColor = &HFF0000
Label3.BackColor = &HFFFF&
Label4.ForeColor = &HFF0000
Label4.BackColor = &HFFFF&
Label5.ForeColor = &HFF0000
Label5.BackColor = &HFFFF&
Label6.ForeColor = &HFF0000
Label6.BackColor = &HFFFF&

End Sub

Private Sub vsbSlide_Change()

    picSlide.Top = -vsbSlide.Value
    picFrame.SetFocus
    
End Sub


Private Sub vsbSlide_Scroll()

    vsbSlide_Change

End Sub

Private Sub picture2_mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Picture3.BackColor = &HFFFF&
Picture4.BackColor = &HFFFF&
Picture5.BackColor = &HFFFF&
Picture6.BackColor = &HFFFF&
Picture7.BackColor = &HFFFF&
Label2.ForeColor = &HFF0000
Label2.BackColor = &HFFFF&
Label3.ForeColor = &HFF0000
Label3.BackColor = &HFFFF&
Label4.ForeColor = &HFF0000
Label4.BackColor = &HFFFF&
Label5.ForeColor = &HFF0000
Label5.BackColor = &HFFFF&
Label6.ForeColor = &HFF0000
Label6.BackColor = &HFFFF&







End Sub






Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HFF0000
Picture3.BackColor = &HFF0000
Label2.ForeColor = &HFFFF&
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.BackColor = &HFF0000
Picture4.BackColor = &HFF0000
Label3.ForeColor = &HFFFF&
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.BackColor = &HFF0000
Picture5.BackColor = &HFF0000
Label4.ForeColor = &HFFFF&
End Sub
Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.BackColor = &HFF0000
Picture6.BackColor = &HFF0000
Label5.ForeColor = &HFFFF&
End Sub


Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.BackColor = &HFF0000
Picture7.BackColor = &HFF0000
Label6.ForeColor = &HFFFF&
End Sub



