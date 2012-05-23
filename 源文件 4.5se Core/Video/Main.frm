VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "S.m.M. Video Capturer"
   ClientHeight    =   3285
   ClientLeft      =   3060
   ClientTop       =   2685
   ClientWidth     =   4590
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   StartUpPosition =   2  '��Ļ����
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuLoadPal 
         Caption         =   "������ɫ��(&L)..."
      End
      Begin VB.Menu mnuSetCapFile 
         Caption         =   "�����ļ�(&S)..."
      End
      Begin VB.Menu mnuAllocFileSpace 
         Caption         =   "�ļ���С(&A)"
      End
      Begin VB.Menu mnuspacer0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveFileAs 
         Caption         =   "��Ƶ���Ϊ(C)..."
      End
      Begin VB.Menu mnuSavePalette 
         Caption         =   "������ɫ��(&P)..."
      End
      Begin VB.Menu mnuSaveFrame 
         Caption         =   "ץȡ��ǰ֡Ϊ(&F)..."
      End
      Begin VB.Menu mnuspacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuCopy 
         Caption         =   "����(C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "ճ����ɫ��(&P)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuspacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "����(&F)..."
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "ѡ��(&O)"
      Begin VB.Menu mnuAudioFmt 
         Caption         =   "��Ƶ��ʽ(&A)..."
      End
      Begin VB.Menu mnuspacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "��ʽ(&F)..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSource 
         Caption         =   "ѡ����Դ(&S)..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "��ʾ(&D)..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuspacer5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompression 
         Caption         =   "ѹ��(C)..."
      End
      Begin VB.Menu mnuspacer6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Ԥ��(&P)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOverlay 
         Caption         =   "����(&O)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuspacer7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDriver 
         Caption         =   "<none>"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuCapture 
      Caption         =   "����(&C)"
      Begin VB.Menu mnuCapFrame 
         Caption         =   "��һ֡(&S)"
      End
      Begin VB.Menu mnuCapFrames 
         Caption         =   "����֡(&F)..."
      End
      Begin VB.Menu mnuCapVid 
         Caption         =   "��Ƶ(&V)..."
      End
      Begin VB.Menu mnuCapPal 
         Caption         =   "��ɫ��(&P)..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private hCapWnd As Long       ' Handle to the Capture Windows
Private nDriverIndex As Long  ' video driver index (default 0)
Private m_CapParams As CAPTUREPARMS
'Public property to prevent reentrancy in Form_Resize event
Public AutoSizing As Boolean
'read-only public property to allow other forms to retrieve hwnd of Cap Window
Public Property Get capwnd() As Long
On Error Resume Next
    capwnd = hCapWnd
End Property
'read-only properties for sizing
Public Property Get MenuHeight() As Long
On Error Resume Next
    MenuHeight = GetSystemMetrics(SM_CYMENU)
End Property
Public Property Get CaptionHeight() As Long
On Error Resume Next
    CaptionHeight = GetSystemMetrics(SM_CYCAPTION)
End Property
Public Property Get XBorder() As Long
On Error Resume Next
    If Me.Appearance = 0 Then   'flat
        XBorder = GetSystemMetrics(SM_CXBORDER)
    Else                        '3D
        XBorder = GetSystemMetrics(SM_CXEDGE)
    End If
End Property
Public Property Get YBorder() As Long
On Error Resume Next
    If Me.Appearance = 0 Then   'flat
        YBorder = GetSystemMetrics(SM_CYBORDER)
    Else                        '3D
        YBorder = GetSystemMetrics(SM_CYEDGE)
    End If
End Property


Private Sub Form_Load()
On Error Resume Next
    If App.PrevInstance Then End

    Dim retVal As Boolean
    Dim numDevs As Long
    
    'load trivial settings first
    Me.BackColor = Val(GetSetting(App.Title, "ѡ��", "����ɫ", "&H404040")) 'default to dk gray
    
    numDevs = VBEnumCapDrivers(Me)
    If 0 = numDevs Then
        MsgBox "�Ҳ�����λ����ͷ������������Ƶ�����豸!", vbCritical, App.Title
       Exit Sub
    End If
    nDriverIndex = Val(GetSetting(App.Title, "Դ", "����", "0"))
    'if invalid entry is in registry use default (0)
    If mnuDriver.UBound < nDriverIndex Then
        nDriverIndex = 0
    End If
    mnuDriver(nDriverIndex).Checked = True
    '//Create Capture Window
    'Call capGetDriverDescription( nDriverIndex,  lpszName, 100, lpszVer, 100  '// Retrieves driver info
    hCapWnd = capCreateCaptureWindow("VB CAP WINDOW", WS_CHILD Or WS_VISIBLE, 0, 0, 160, 120, Me.hWnd, 0)
    If 0 = hCapWnd Then
        MsgBox "�������񴰿�ʧ��!", vbCritical, App.Title
        Exit Sub
    End If
    retVal = ConnectCapDriver(hCapWnd, nDriverIndex)
    If False = retVal Then
        MsgBox "������λ�����豸ʱ��������!", vbInformation, App.Title
    Else
        #If USECALLBACKS = 1 Then
            ' if we have a valid capwnd we can enable our status callback function
            Call capSetCallbackOnStatus(hCapWnd, AddressOf StatusProc)
            Debug.Print " "
        #End If
    End If
        '// Set the video stream callback function
'    capSetCallbackOnVideoStream lwndC, AddressOf MyVideoStreamCallback
'    capSetCallbackOnFrame lwndC, AddressOf MyFrameCallback
 

End Sub


Public Sub Form_Resize()
    On Error Resume Next
    Dim retVal As Boolean
    Dim capStat As CAPSTATUS
    'kludgy way to restrict min form size - better way is to subclass MINMAXINFO messages
    If Me.ScaleWidth < 320 Then Me.Width = (320 + (Me.XBorder * 2)) * Screen.TwipsPerPixelX
    If Me.ScaleHeight < 240 Then Me.Height = (240 + (Me.YBorder * 2) + Me.MenuHeight + Me.CaptionHeight) * Screen.TwipsPerPixelY
    'Get the capture window attributes
    retVal = capGetStatus(hCapWnd, capStat)
        
    If retVal Then
        'center the capture window on the form
        Call SetWindowPos(hCapWnd, _
                    0&, _
                    (Me.ScaleWidth - capStat.uiImageWidth) / 2, _
                    (Me.ScaleHeight - capStat.uiImageHeight) / 2, _
                    0&, _
                    0&, _
                    SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOSENDCHANGING) 'by telling Windows not to send
                                                                    'WM_WINDOWPOSCHANGING messages we
                                                                    'eliminate the need for a reentrancy flag
    End If
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
#If USECALLBACKS = 1 Then
    ' Disable status callback
    Call capSetCallbackOnStatus(hCapWnd, 0&)
    Debug.Print " "
#End If

'disconnect VFW driver
Call mVFW.capDriverDisconnect(hCapWnd)
'destroy CapWnd
If hCapWnd <> 0 Then Call DestroyWindow(hCapWnd)

End Sub


'Private Sub mnuAbout_Click()
'On Error Resume Next
'    Dim AboutWnd As frmAbout
'    Set AboutWnd = New frmAbout
'
'    AboutWnd.Show vbModal, Me
'
'    Set AboutWnd = Nothing
'End Sub

Private Sub mnuAllocFileSpace_Click()
On Error Resume Next
    Dim AllocWnd As frmAlloc
    Set AllocWnd = New frmAlloc
    
    AllocWnd.Show vbModal, Me
    
    Set AllocWnd = Nothing

End Sub

Private Sub mnuAudioFmt_Click()
On Error Resume Next
    Call SetAudioFormatDlg(Me.hWnd)
End Sub

Private Sub mnuCapFrame_Click()
On Error Resume Next
    Call capGrabFrame(hCapWnd)

End Sub

Private Sub mnuCapFrames_Click()
On Error Resume Next
    Dim FrameCapWnd As frmCapFrame
    
    Set FrameCapWnd = New frmCapFrame
    FrameCapWnd.Show vbModal, Me
    
    Set FrameCapWnd = Nothing
    
End Sub

Private Sub mnuCapPal_Click()
On Error Resume Next
    Dim PalCapWnd As frmCapPal
    
    Set PalCapWnd = New frmCapPal
    PalCapWnd.Show vbModal, Me
    
    Set PalCapWnd = Nothing
End Sub

Private Sub mnuCapVid_Click()
On Error Resume Next
    Dim retVal As Boolean
    Dim VidCapWnd As frmCapVid
    
    Set VidCapWnd = New frmCapVid
    VidCapWnd.Show vbModal, Me
    If VidCapWnd.Tag <> "" Then 'use tag to indicate whether user has pressed OK or not
'            // Capture video sequence
        retVal = capCaptureSequence(hCapWnd)
        Unload VidCapWnd 'reclaim mem
    End If
    Set VidCapWnd = Nothing
End Sub

Private Sub mnuCompression_Click()
On Error Resume Next
    Call capDlgVideoCompression(hCapWnd)

End Sub

Private Sub mnuCopy_Click()
    On Error Resume Next
    Call capEditCopy(hCapWnd)

End Sub

Private Sub mnuDisplay_Click()
On Error Resume Next
    Call capDlgVideoDisplay(hCapWnd)
    
End Sub

Private Sub mnuDriver_Click(Index As Integer)
On Error Resume Next
    Dim retVal As Boolean
    
    retVal = ConnectCapDriver(hCapWnd, Index)
    If False = retVal Then
        MsgBox "������Ƶ�����豸�����������ʱ��������!", vbInformation, App.Title
    Else
        Call SaveSetting(App.Title, "Դ", "����", CStr(Index)) 'save selected device index
    End If
End Sub

Private Sub mnuExit_Click()
On Error Resume Next
    Unload Me
    
End Sub

Private Sub mnuFormat_Click()
On Error Resume Next
    Call capDlgVideoFormat(hCapWnd)
    Call ResizeCaptureWindow(hCapWnd)

End Sub

Private Sub mnuLoadPal_Click()
On Error Resume Next
Dim PalFile As String
Dim PalFileTitle As String
Dim retVal As Boolean

retVal = VBGetOpenFileName(PalFile, _
                            PalFileTitle, _
                            filter:="��ɫ���ļ� (*.pal)|*.pal", _
                            InitDir:=App.path, _
                            DlgTitle:="���ص�ɫ��", _
                            DefaultExt:="��ɫ��", _
                            HideReadOnly:=True, _
                            Owner:=Me.hWnd)
If True = retVal Then 'user did not cancel
    retVal = capPaletteOpen(hCapWnd, PalFile)
    If 0 = retVal Then
        MsgBox "���ܼ��ص�ɫ���ļ�: " & PalFileTitle, vbInformation, App.Title
    End If
End If
        

End Sub

Private Sub mnuOverlay_Click()
    On Error Resume Next
    mnuOverlay.Checked = Not (mnuOverlay.Checked)
    Call capOverlay(hCapWnd, mnuOverlay.Checked)
    
End Sub

Private Sub mnuPreferences_Click()
On Error Resume Next
    Dim PrefsWnd As frmPrefs
    
    Set PrefsWnd = New frmPrefs
    PrefsWnd.Show vbModal, Me
    
    Set PrefsWnd = Nothing
End Sub

Private Sub mnuPreview_Click()
On Error Resume Next
    mnuPreview.Checked = Not (mnuPreview.Checked)
    Call capPreview(hCapWnd, mnuPreview.Checked)

End Sub


Private Sub mnuSaveFileAs_Click()
On Error Resume Next
Dim FileName As String
Dim retVal As Boolean

retVal = VBGetSaveFileNamePreview(FileName, _
                            FileMustExist:=False, _
                            HideReadOnly:=True, _
                            filter:="AVI �ļ� (*.avi)|*.avi", _
                            DefaultExt:="avi", _
                            Owner:=Me.hWnd)
If False <> retVal Then
    retVal = capFileSaveAs(hCapWnd, FileName)
    If True <> retVal Then
        MsgBox "�����ļ��Ƿ�������!", vbInformation, App.Title
    End If
End If
End Sub

Private Sub mnuSaveFrame_Click()
On Error Resume Next
Dim FileName As String
Dim retVal As Boolean

retVal = VBGetSaveFileName(FileName, _
                            filter:="λͼ�ļ� (*.bmp)|*.bmp", _
                            DlgTitle:="���浥һ֡", _
                            DefaultExt:="bmp", _
                            Owner:=Me.hWnd)
If False <> retVal Then
    retVal = capFileSaveDIB(hCapWnd, FileName)
    If True <> retVal Then
        MsgBox "���浥һ֡ʱ��������!", vbInformation, App.Title
    End If
End If
End Sub

Private Sub mnuSavePalette_Click()
On Error Resume Next
Dim FileName As String
Dim retVal As Boolean

retVal = VBGetSaveFileName(FileName, _
                            filter:="��ɫ���ļ� (*.pal)|*.pal", _
                            DlgTitle:="�����ɫ��", _
                            DefaultExt:="pal", _
                            Owner:=Me.hWnd)
If False <> retVal Then
    retVal = capPaletteSave(hCapWnd, FileName)
    If True <> retVal Then
        MsgBox "�����ɫ��ʱ��������!", vbInformation, App.Title
    End If
End If
End Sub

Private Sub mnuSetCapFile_Click()
On Error Resume Next
Dim CapFile As String
Dim CapFileTitle As String
Dim CapFileDir As String
Dim retVal As Boolean
Dim nfileLen As Long

CapFile = mVFW.capFileGetCaptureFile(hCapWnd)
CapFileTitle = VBGetFileTitle(CapFile)
CapFileDir = Left$(CapFile, Len(CapFile) - Len(CapFileTitle))
retVal = VBGetOpenFileNamePreview(CapFile, _
                            FileTitle:=CapFileTitle, _
                            filter:="AVI �ļ� (*.avi)|*.avi", _
                            InitDir:=CapFileDir, _
                            DlgTitle:="���ò����ļ�", _
                            FileMustExist:=False, _
                            HideReadOnly:=True, _
                            DefaultExt:="avi", _
                            Owner:=Me.hWnd)
If True = retVal Then 'user did not cancel
    retVal = mVFW.capFileSetCaptureFile(hCapWnd, CapFile)
    If 0 = retVal Then
        MsgBox "��������Ƶ�ļ�ʧ��: " & CapFileTitle, vbInformation, App.Title
        Exit Sub
    Else
        'capture file was changed successfully let's allocate some disk space for it
        'but only if it doesn't already exist
        On Error Resume Next
        nfileLen = FileLen(CapFile)
        If Err.Number = 53 Then 'file does not exist
            Call mnuAllocFileSpace_Click
        End If
    End If
End If
End Sub

Private Sub mnuSource_Click()
'   /*
'    * Display the Video Source dialog when "Source" is selected from the
'    * menu bar.
'    */
    On Error Resume Next
    Call capDlgVideoSource(hCapWnd)

End Sub


