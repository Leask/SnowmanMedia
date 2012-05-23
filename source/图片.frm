VERSION 5.00
Begin VB.MDIForm mfrmMain 
   BackColor       =   &H00000000&
   Caption         =   "Snowman Media Pictures Browser  1.0"
   ClientHeight    =   5025
   ClientLeft      =   1590
   ClientTop       =   1875
   ClientWidth     =   6360
   Icon            =   "图片.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   6360
      TabIndex        =   0
      Top             =   4695
      Width           =   6360
      Begin SmMPicturesBrowser10.ProgressBar prgMain 
         Height          =   285
         Left            =   0
         Top             =   45
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   503
         ForeColor       =   16711680
         BackColor       =   65535
         Min             =   1
      End
      Begin VB.Label lblSize 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6660
         TabIndex        =   3
         Top             =   45
         Width           =   1545
      End
      Begin VB.Label lblImage 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5085
         TabIndex        =   2
         Top             =   45
         Width           =   1545
      End
      Begin VB.Label lblStatus 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ready."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   60
         Width           =   5055
      End
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFile 
         Caption         =   "浏览(&B)"
         Index           =   0
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "打开(&O)"
         Index           =   2
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "打开动画Gif(&G)"
         Index           =   3
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFile 
         Caption         =   "幻灯片方式(&M)"
         Index           =   4
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuFile 
         Caption         =   "保存(&S)"
         Index           =   5
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "打印(&P)"
         Index           =   7
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   10
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   11
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   12
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuFile 
         Caption         =   "退出(&X)"
         Index           =   14
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEditTOP 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit 
         Caption         =   "复制(&C)"
         Index           =   1
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "粘贴(&V)"
         Index           =   2
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuImageTOP 
      Caption         =   "图象(&I)"
      Begin VB.Menu mnuImage 
         Caption         =   "柔化..."
         Index           =   0
         Begin VB.Menu mnuLowPass 
            Caption         =   "淡化"
            Index           =   0
         End
         Begin VB.Menu mnuLowPass 
            Caption         =   "深层淡化"
            Index           =   1
         End
         Begin VB.Menu mnuLowPass 
            Caption         =   "模糊"
            Index           =   2
         End
         Begin VB.Menu mnuLowPass 
            Caption         =   "深层模糊"
            Index           =   3
         End
      End
      Begin VB.Menu mnuImage 
         Caption         =   "锐化..."
         Index           =   1
         Begin VB.Menu mnuHighPass 
            Caption         =   "锐化"
            Index           =   0
         End
         Begin VB.Menu mnuHighPass 
            Caption         =   "深层锐化"
            Index           =   1
         End
         Begin VB.Menu mnuHighPass 
            Caption         =   "钝化"
            Index           =   2
         End
      End
      Begin VB.Menu mnuImage 
         Caption         =   "预定义滤镜..."
         Index           =   2
         Begin VB.Menu mnuSpecial 
            Caption         =   "浮雕"
            Index           =   0
         End
         Begin VB.Menu mnuSpecial 
            Caption         =   "杂点..."
            Index           =   2
         End
         Begin VB.Menu mnuSpecial 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuSpecial 
            Caption         =   "初级恶化"
            Index           =   4
         End
         Begin VB.Menu mnuSpecial 
            Caption         =   "中级恶化"
            Index           =   5
         End
         Begin VB.Menu mnuSpecial 
            Caption         =   "高级恶化"
            Index           =   6
         End
      End
      Begin VB.Menu mnuImage 
         Caption         =   "自定义滤镜..."
         Index           =   3
      End
      Begin VB.Menu mnuImage 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuImage 
         Caption         =   "改变大小..."
         Index           =   5
      End
      Begin VB.Menu mnuImage 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuImage 
         Caption         =   "合并图片..."
         Index           =   7
      End
   End
   Begin VB.Menu mnuColorTOP 
      Caption         =   "颜色(&C)"
      Begin VB.Menu mnuColors 
         Caption         =   "暗淡"
         Index           =   0
      End
      Begin VB.Menu mnuColors 
         Caption         =   "发亮"
         Index           =   1
      End
      Begin VB.Menu mnuColors 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuColors 
         Caption         =   "渗透..."
         Index           =   3
      End
      Begin VB.Menu mnuColors 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuColors 
         Caption         =   "反色"
         Index           =   5
      End
      Begin VB.Menu mnuColors 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuColors 
         Caption         =   "灰度"
         Index           =   7
      End
      Begin VB.Menu mnuColors 
         Caption         =   "黑白"
         Index           =   8
      End
      Begin VB.Menu mnuColors 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuColors 
         Caption         =   "调色板..."
         Index           =   11
      End
   End
   Begin VB.Menu hqTop 
      Caption         =   "获取(&R)"
      Begin VB.Menu hqs 
         Caption         =   "屏幕获取"
      End
      Begin VB.Menu hqd 
         Caption         =   "扫描获取"
      End
      Begin VB.Menu hq 
         Caption         =   "视频获取..."
         Begin VB.Menu mnuStart 
            Caption         =   "开始"
         End
         Begin VB.Menu mnuAllocate 
            Caption         =   "分配"
         End
         Begin VB.Menu mnuhq 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCopy 
            Caption         =   "复制"
         End
         Begin VB.Menu mnuPreview 
            Caption         =   "预览"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuhql 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDisplay 
            Caption         =   "显示"
         End
         Begin VB.Menu mnuFormat 
            Caption         =   "格式"
         End
         Begin VB.Menu mnuScale 
            Caption         =   "标尺"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuhqa 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSource 
            Caption         =   "资源"
         End
         Begin VB.Menu mnuSelect 
            Caption         =   "选择..."
         End
         Begin VB.Menu mnuCompression 
            Caption         =   "压缩"
         End
      End
   End
   Begin VB.Menu mnuWindowTop 
      Caption         =   "窗口(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindow 
         Caption         =   "水平平铺(&F&1)"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "垂直平铺(&F&2)"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "层叠窗口(&F&3)"
         Index           =   2
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuWindow 
         Caption         =   "重排图标(&F&4)"
         Index           =   3
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu Help 
      Caption         =   "辅助(&H)"
      Begin VB.Menu Zoon 
         Caption         =   "屏幕放大镜(&Z)"
         Shortcut        =   ^Z
      End
   End
End
Attribute VB_Name = "mfrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_cMRU As New cMRUFileList
Private m_bInIDE As Boolean
Private m_lCount As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long
Public Enum EShellShowConstants
    essSW_HIDE = 0
    essSW_MAXIMIZE = 3
    essSW_MINIMIZE = 6
    essSW_SHOWMAXIMIZED = 3
    essSW_SHOWMINIMIZED = 2
    essSW_SHOWNORMAL = 1
    essSW_SHOWNOACTIVATE = 4
    essSW_SHOWNA = 8
    essSW_SHOWMINNOACTIVE = 7
    essSW_SHOWDEFAULT = 10
    essSW_RESTORE = 9
    essSW_SHOW = 5
End Enum
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5        ' access denied
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2                ' file not found
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_PNF = 3                ' path not found
Private Const SE_ERR_OOM = 8                ' out of memory
Private Const SE_ERR_SHARE = 26


Private Const MAX_PATH = 260

Public Property Get NewImageIndex() As Long
   m_lCount = m_lCount + 1
   NewImageIndex = m_lCount
End Property

Public Function ShellEx( _
        ByVal sFIle As String, _
        Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, _
        Optional ByVal sParameters As String = "", _
        Optional ByVal sDefaultDir As String = "", _
        Optional sOperation As String = "open", _
        Optional Owner As Long = 0 _
    ) As Boolean
Dim lR As Long
Dim lErr As Long, sErr As Long
    If (InStr(UCase$(sFIle), ".EXE") <> 0) Then
        eShowCmd = 0
    End If
    On Error Resume Next
    If (sParameters = "") And (sDefaultDir = "") Then
        lR = ShellExecuteForExplore(Owner, sOperation, sFIle, 0, 0, essSW_SHOWNORMAL)
    Else
        lR = ShellExecute(Owner, sOperation, sFIle, sParameters, sDefaultDir, eShowCmd)
    End If
    If (lR < 0) Or (lR > 32) Then
        ShellEx = True
    Else
        ' raise an appropriate error:
        lErr = vbObjectError + 1048 + lR
        Select Case lR
        Case 0
            lErr = 7: sErr = "内存溢出."
        Case ERROR_FILE_NOT_FOUND
            lErr = 53: sErr = "找不到指定文件."
        Case ERROR_PATH_NOT_FOUND
            lErr = 76: sErr = "找不到指定路径."
        Case ERROR_BAD_FORMAT
            sErr = "找不到指定的执行文件或其文件出错."
        Case SE_ERR_ACCESSDENIED
            lErr = 75: sErr = "路径或文件数据错误."
        Case SE_ERR_ASSOCINCOMPLETE
            sErr = "这种类型的文件没有指定合适的打开方式."
        Case SE_ERR_DDEBUSY
            lErr = 285: sErr = "文件由于正在被其他应用程序使用而无法打开,请稍后再试."
        Case SE_ERR_DDEFAIL
            lErr = 285: sErr = "文件由于数据交换错误而无法打开,请稍后再试."
        Case SE_ERR_DDETIMEOUT
            lErr = 286: sErr = "文件打开超时,请稍后再试"
        Case SE_ERR_DLLNOTFOUND
            lErr = 48: sErr = "找不到特定的数据引擎."
        Case SE_ERR_FNF
            lErr = 53: sErr = "找不到指定文件."
        Case SE_ERR_NOASSOC
            sErr = "这种类型的文件没有指定合适的打开方式."
        Case SE_ERR_OOM
            lErr = 7: sErr = "内存溢出."
        Case SE_ERR_PNF
            lErr = 76: sErr = "找不到指定路径."
        Case SE_ERR_SHARE
            lErr = 75: sErr = "发生意外错误."
        Case Else
            sErr = "Snowman Media Pictures Browser  1.0 在打开或打印图片时发生错误."
        End Select
                
        Err.Raise lErr, , App.EXEName & ".GShell", sErr
        ShellEx = False
    End If

End Function

Public Property Get TempDir() As String
Dim sRet As String, c As Long
    sRet = String$(MAX_PATH, 0)
    c = GetTempPath(MAX_PATH, sRet)
    If c = 0 Then Err.Raise Err.LastDllError
    TempDir = Left$(sRet, c)
End Property
Public Property Get TempFileName( _
        Optional ByVal sPrefix As String, _
        Optional ByVal sPathName As String) As String
Dim iPos As Long
    If sPrefix = "" Then sPrefix = ""
    If sPathName = "" Then sPathName = TempDir
    
    Dim sRet As String
    sRet = String(MAX_PATH, 0)
    GetTempFileName sPathName, sPrefix, 0, sRet
    If (Err.LastDllError <> 0) Then Err.Raise Err.LastDllError
    iPos = InStr(sRet, Chr$(0))
    If (iPos <> 0) Then
        TempFileName = Left$(sRet, (iPos - 1))
    Else
        TempFileName = sRet
    End If
End Property

Private Function InIDECheck() As Boolean
    m_bInIDE = True
    InIDECheck = True
End Function

Public Sub AddMRUFile(ByVal sFIle As String)
    m_cMRU.AddFile sFIle
    pShowMRU
End Sub
Public Property Let ProgressMax(ByVal lMax As Long)
    prgMain.Max = lMax
End Property
Public Property Let ProgressValue(ByVal lValue As Long)
    prgMain.Position = lValue
End Property
Public Property Let ShowProgress(ByVal bShow As Boolean)
    prgMain.Visible = bShow
End Property

Public Sub SetStatus( _
        Optional ByVal sMain As String = "#", _
        Optional ByVal sImage As String = "#", _
        Optional ByVal sSize As String = "#" _
    )
    If (sMain <> "#") Then
        lblStatus.Caption = " " & sMain
    End If
    If (sImage <> "#") Then
        lblImage.Caption = " " & sImage
    End If
    If (sSize <> "#") Then
        lblSize.Caption = " " & sSize
    End If
End Sub

Private Function GetActiveform(ByRef f As frmImage) As Boolean
    If Not (Me.ActiveForm Is Nothing) Then
        If (Me.ActiveForm.Name = "frmImage") Then
            Set f = Me.ActiveForm
            GetActiveform = True
        Else
            MsgBox "请先选择一幅图片再进行加工.", vbInformation
        End If
    Else
        MsgBox "请先选择一幅图片再进行加工.", vbInformation
    End If
End Function

Private Sub pOpen(Optional ByVal sFIle As String = "")
Dim c As New GCommonDialog
Dim bContinue As Boolean
    
    bContinue = True
    If (sFIle = "") Then
        ' Get a new file:
        bContinue = False
        If (c.VBGetOpenFileName(sFIle, , , , , , "图片文件(*.BMP;*.GIF;*.JPG;*.DIB)|*.BMP;*.GIF;*.JPG;*.DIB|位图文件(*.BMP;*.DIB)|*.BMP;*.DIB|Gig文件(*.GIF)|*.GIF|Jpeg文件(*.JPG)|*.JPG|所有文件(*.*)|*.*", 1, , , "BMP", Me.hWnd)) Then
            bContinue = True
        End If
    End If
    
    If (bContinue) Then
        Dim f As New frmImage
        If (f.OpenFile(sFIle)) Then
            f.Show
        Else
            Unload f
        End If
    End If
End Sub

Private Sub pSave()
Dim f As frmImage
    If (GetActiveform(f)) Then
        f.SaveFile
    End If
End Sub
Private Sub pShowMRU()
Dim i As Long
    For i = 1 To m_cMRU.FileCount
        If (m_cMRU.FileExists(i)) Then
            mnuFile(i + 8).Visible = True
            mnuFile(i + 8).Caption = m_cMRU.MenuCaption(i)
        End If
    Next i
    mnuFile(13).Visible = (m_cMRU.FileCount > 0)
End Sub



Private Sub hqd_Click()
 Form3.Show
End Sub


Private Sub hqs_Click()
Formjp.Show
Me.Hide
End Sub

Private Sub MDIForm_Load()
   

Dim cR As New cRegistry
Dim lHDC As Long
Dim lhWNd As Long
Dim sMsg As String


    m_cMRU.MaxFileCount = 4
    cR.ClassKey = HKEY_CURRENT_USER
    cR.SectionKey = "Software\vbAccelerator\vbImageProc"
    m_cMRU.Load cR
    pShowMRU
    Me.Show
    Debug.Assert (InIDECheck = True)
   ' If (m_bInIDE) Then
    '    MsgBox "You are running this sample in the VB IDE." & vbCrLf & vbCrLf & "Please note that the Image Processing functions run 25 - 50x quicker when compiled to Native Code.", vbInformation
    'End If

    lhWNd = GetDesktopWindow()
    lHDC = GetDC(lhWNd)
   ' If (GetDeviceCaps(lHDC, BITSPIXEL) <= 8) Then
    '    sMsg = "Screen colour depths below 16 bits/pixel are not supported by this sample."
     '   If (m_bInIDE) Then
      '      sMsg = sMsg & vbCrLf & vbCrLf & "You must exit out of VB, change colour depth and re-load in VB to get it to work."
       ' End If
       ' MsgBox sMsg, vbExclamation
    'End If
    ReleaseDC lhWNd, lHDC
Dim f As New frmThumbs
     f.Show


 








End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Long
    If UnloadMode <> vbAppWindows And UnloadMode <> vbAppTaskManager Then
        For i = 0 To Forms.Count - 1
            If (Forms(i).Name = "frmImage") Then
                If (Forms(i).Dirty) Then
                    If Not (Forms(i).QuerySave()) Then
                        Cancel = True
                        Exit Sub
                    End If
                End If
            End If
        Next i
    End If
    
    Dim cR As New cRegistry
    cR.ClassKey = HKEY_CURRENT_USER
    cR.SectionKey = "Software\vbAccelerator\vbImageProc"
    m_cMRU.Save cR
   
   
    
    
    
    
    
  capSetCallbackOnError lwndC, vbNull
    capSetCallbackOnStatus lwndC, vbNull
    capSetCallbackOnYield lwndC, vbNull
    capSetCallbackOnFrame lwndC, vbNull
    capSetCallbackOnVideoStream lwndC, vbNull
    capSetCallbackOnWaveStream lwndC, vbNull
    capSetCallbackOnCapControl lwndC, vbNull
   End
End Sub

Private Sub mnuColors_Click(Index As Integer)
Dim f As frmImage
    If (GetActiveform(f)) Then
        Select Case Index
        Case 0
            f.Fade
        Case 1
            f.Lighten
        Case 3
            pColourise f
        Case 5
            f.NegativeImage
        Case 7
            f.GrayScale
        Case 8
            f.BlackAndWhite
        Case 11
            pPalette f
        End Select
    End If
End Sub

Private Sub mnuEdit_Click(Index As Integer)
Dim f As frmImage
Dim sName As String
    Select Case Index
    Case 1
        If (GetActiveform(f)) Then
            f.CopyImage
        End If
    Case 2
        On Error GoTo PasteImageError
        Dim sPic As New StdPicture
        Set sPic = Clipboard.GetData(vbCFBitmap)
        sName = TempFileName("VBIM")
        SavePicture sPic, sName
        Dim fN As New frmImage
        If (fN.OpenFile(sName, True)) Then
            fN.Show
        Else
            Unload fN
        End If
        On Error Resume Next
        Kill sName

    End Select
    Exit Sub
PasteImageError:
    MsgBox "Snowman Media Pictures Browser  1.0 在粘贴文件时发生错误: " & Err.Description, vbExclamation
    On Error Resume Next
    Kill sName
    Exit Sub
      Resume 0
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Select Case Index
    Case 0
     Dim f As New frmThumbs
     f.Show
    Case 2
        pOpen
    Case 3
         Dim fm As New FrmAniGif
        fm.Show
    Case 4
        Form1.Show
        Form2.Show
    Case 5
        pSave
    Case 7
        MsgBox "Snowman Media Pictures Browser  1.0 在寻找打印机时发生错误.", vbInformation
    Case 9 To 12
        pOpen m_cMRU.file(Index - 8)
    Case 14
        Unload Me
    End Select
End Sub



Private Sub mnuHighPass_Click(Index As Integer)
Dim f As frmImage
    If (GetActiveform(f)) Then
        Select Case Index
        Case 0
            f.ProcessImage eSharpen
        Case 1
            f.ProcessImage eSharpenMore
        Case 2
            f.ProcessImage eUnSharp
        End Select
    End If
End Sub

Private Sub mnuImage_Click(Index As Integer)
Dim f As frmImage
   Select Case Index
   Case 3
      ' User defined filter...
      If (GetActiveform(f)) Then
          pCustomFilter f
      End If
   Case 5
      ' Resample....
      If (GetActiveform(f)) Then
          pResample f
      End If
   Case 7
      ' Combine:
      If (GetActiveform(f)) Then
         pCombine
      End If
   End Select
End Sub
Private Function pResample(ByRef f As frmImage) As Boolean
    Dim fC As New frmNewSize
    fC.SetSize f.ImageWidth, f.ImageHeight
    fC.Show vbModal, Me
    If Not (fC.Cancelled) Then
        f.Resample fC.ImageWidth, fC.ImageHeight
        pResample = True
    End If
End Function
Private Function pCustomFilter(ByRef f As frmImage) As Boolean
    Dim fC As New frmCustomFilter
    fC.Show vbModal, Me
    If Not (fC.Cancelled) Then
        f.LoadCustomFilter fC.ImageProcess
        f.ProcessImage eCustom
        pCustomFilter = True
    End If
End Function
Private Function pCombine() As Boolean
   Dim fC As New frmCombination
   fC.Show vbModal, Me
   If Not (fC.Cancelled) Then
      Dim f As New frmImage
      f.Show
      f.Combine fC
   End If
End Function

Private Sub mnuLowPass_Click(Index As Integer)
Dim f As frmImage
    If (GetActiveform(f)) Then
        Select Case Index
        Case 0
            f.ProcessImage eSoften
        Case 1
            f.ProcessImage eSoftenMore
        Case 2
            f.ProcessImage eBlur
        Case 3
            f.ProcessImage eBlurMore
        End Select
    End If
End Sub

Private Sub mnuPreview_Click()
 'frmMain.StatusBar.SimpleText = vbNullString
    mnuPreview.Checked = Not (mnuPreview.Checked)
    capPreview lwndC, mnuPreview.Checked
End Sub

Private Sub mnuSpecial_Click(Index As Integer)
Dim f As frmImage
    If (GetActiveform(f)) Then
        Select Case Index
        Case 0
            ' Emboss:
            f.ProcessImage eEmboss
        Case 2
            ' Add noise:
            pAddNoise f
        Case 4
            ' Minimum:
            f.ProcessImage eMinimum
        Case 5
            ' Median:
            f.ProcessImage eMedian
        Case 6
            ' Maximum:
            f.ProcessImage eMaximum
        End Select
    End If
End Sub
Private Sub pAddNoise(ByRef f As frmImage)
Dim fC As New frmAddNoise
    fC.Show vbModal, Me
    If Not (fC.Cancelled) Then
        f.AddNoise fC.Random, fC.Percentage
    End If
End Sub

Private Sub pColourise(ByRef f As frmImage)
Dim fC As New frmColourise
   fC.Show vbModal, Me
   If Not (fC.Cancelled) Then
      f.Colourise fC.Hue
   End If
End Sub

Private Sub pPalette(ByRef f As frmImage)
Dim fC As New frmPalette
   fC.Show vbModal, Me
   If Not (fC.Cancelled) Then
      f.ApplyPalette fC.FileName
   End If
End Sub

Private Sub mnuWindow_Click(Index As Integer)
    Select Case Index
    Case 0
        Me.Arrange vbTileHorizontal
    Case 1
        Me.Arrange vbTileVertical
    Case 2
        Me.Arrange vbCascade
    Case 3
        Me.Arrange vbArrangeIcons
    End Select
End Sub


Private Sub mnuAllocate_Click()

 Dim sFIle As String * 250
 Dim lSize As Long
 
 '// Setup swap file for capture
 lSize = 1000000
 sFIle = "C:\TEMP.AVI"
 capFileSetCaptureFile lwndC, sFIle
 capFileAlloc lwndC, lSize
 
End Sub


Private Sub mnuCompression_Click()
'   /*
'   * Display the Compression dialog when "Compression" is selected from
'   * the menu bar.
'   */
    
    capDlgVideoCompression lwndC

End Sub

Private Sub mnuCopy_Click()

    capEditCopy lwndC
        
End Sub

Private Sub mnuDisplay_Click()
'   /*
'   * Display the Video Display dialog when "Display" is selected from
'   * the menu bar.
'   */

    capDlgVideoDisplay lwndC
    
End Sub



Private Sub mnuFormat_Click()
'  /*
'   * Display the Video Format dialog when "Format" is selected from the
'   * menu bar.
'   */

    capDlgVideoFormat lwndC
    ResizeCaptureWindow lwndC

End Sub



Private Sub mnuScale_Click()
    
    mnuScale.Checked = Not (mnuScale.Checked)
    capPreviewScale lwndC, mnuScale.Checked
    
    If mnuScale.Checked Then
       SetWindowLong lwndC, GWL_STYLE, WS_THICKFRAME Or WS_CAPTION Or WS_VISIBLE Or WS_CHILD
    Else
       SetWindowLong lwndC, GWL_STYLE, WS_BORDER Or WS_CAPTION Or WS_VISIBLE Or WS_CHILD
    End If

    ResizeCaptureWindow lwndC
    
End Sub

Private Sub mnuSelect_Click()
    
    frmSelect.Show vbModal, Me

End Sub

Private Sub mnuSource_Click()
'   /*
'    * Display the Video Source dialog when "Source" is selected from the
'    * menu bar.
'    */
    
    capDlgVideoSource lwndC

End Sub

Private Sub mnuStart_Click()
' /*
'  * If Start is selected from the menu, start Streaming capture.
'  * The streaming capture is terminated when the Escape key is pressed
'  */
    
    Dim lpszName As String * 100
    Dim lpszVer As String * 100
    Dim Caps As CAPDRIVERCAPS
        
    '//Create Capture Window
    capGetDriverDescriptionA 0, lpszName, 100, lpszVer, 100  '// Retrieves driver info
    lwndC = capCreateCaptureWindowA(lpszName, WS_CAPTION Or WS_THICKFRAME Or WS_VISIBLE Or WS_CHILD, 0, 0, 160, 120, Me.hWnd, 0)

    '// Set title of window to name of driver
    SetWindowText lwndC, lpszName
    
    '// Set the video stream callback function
    capSetCallbackOnStatus lwndC, AddressOf MyStatusCallback
    capSetCallbackOnError lwndC, AddressOf MyErrorCallback
    
    '// Connect the capture window to the driver
    If capDriverConnect(lwndC, 0) Then
        '/////
        '// Only do the following if the connect was successful.
        '// if it fails, the error will be reported in the call
        '// back function.
        '/////
        '// Get the capabilities of the capture driver
        capDriverGetCaps lwndC, VarPtr(Caps), Len(Caps)
        
        '// If the capture driver does not support a dialog, grey it out
        '// in the menu bar.
        If Caps.fHasDlgVideoSource = 0 Then mnuSource.Enabled = False
        If Caps.fHasDlgVideoFormat = 0 Then mnuFormat.Enabled = False
        If Caps.fHasDlgVideoDisplay = 0 Then mnuDisplay.Enabled = False
        
        '// Turn Scale on
        capPreviewScale lwndC, True
            
        '// Set the preview rate in milliseconds
        capPreviewRate lwndC, 66
        
        '// Start previewing the image from the camera
        capPreview lwndC, True
            
        '// Resize the capture window to show the whole image
        ResizeCaptureWindow lwndC

    End If



    
    
    
    
    
    
    
    
    Dim sFileName As String
    Dim CAP_PARAMS As CAPTUREPARMS
    
    capCaptureGetSetup lwndC, VarPtr(CAP_PARAMS), Len(CAP_PARAMS)
    
    CAP_PARAMS.dwRequestMicroSecPerFrame = (1 * (10 ^ 6)) / 30  ' 30 Frames per second
    CAP_PARAMS.fMakeUserHitOKToCapture = True
    CAP_PARAMS.fCaptureAudio = False
    
    capCaptureSetSetup lwndC, VarPtr(CAP_PARAMS), Len(CAP_PARAMS)
    
    sFileName = "C:\myvideo.avi"
    
    capCaptureSequence lwndC  ' Start Capturing!
    capFileSaveAs lwndC, sFileName  ' Copy video from swap file into a real file.

End Sub







Private Sub picStatus_Resize()
Dim lW As Long
    On Error Resume Next
    lW = lblImage.Width + 2 * Screen.TwipsPerPixelX + lblSize.Width + 2 * Screen.TwipsPerPixelX
    If (Me.ScaleWidth - lW < 64 * Screen.TwipsPerPixelX) Then
        lblStatus.Width = Me.ScaleWidth - lblStatus.Left * 2
        prgMain.Width = lblStatus.Width
        lblSize.Visible = False
        lblImage.Visible = False
    Else
        lblSize.Visible = True
        lblImage.Visible = True
        lblStatus.Width = Me.ScaleWidth - lblStatus.Left * 2 - lW
        prgMain.Width = lblStatus.Width
        lblImage.Left = lblStatus.Left * 2 + lblStatus.Width + 2 * Screen.TwipsPerPixelX
        lblSize.Left = lblImage.Left + lblImage.Width + 2 * Screen.TwipsPerPixelX
    End If
End Sub






Private Sub Zoon_Click()
frmMainB.Show
End Sub
