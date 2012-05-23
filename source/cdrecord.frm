VERSION 5.00
Begin VB.Form Form101 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CD to Wave  - Snowman Media  2.0"
   ClientHeight    =   5055
   ClientLeft      =   4845
   ClientTop       =   1860
   ClientWidth     =   7365
   Icon            =   "cdrecord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5055
   ScaleWidth      =   7365
   Begin VB.FileListBox File1 
      Height          =   1890
      Left            =   1080
      TabIndex        =   19
      Top             =   6300
      Width           =   2580
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "开始(&S)"
      Height          =   375
      Left            =   45
      TabIndex        =   5
      Top             =   4320
      Width           =   2265
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "停止(&T)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2430
      TabIndex        =   6
      Top             =   4320
      Width           =   2265
   End
   Begin VB.CommandButton cmdFrom 
      Caption         =   "<<"
      Height          =   375
      Left            =   3555
      TabIndex        =   17
      Top             =   3960
      Width           =   1140
   End
   Begin VB.CommandButton cmdTo 
      Caption         =   ">>"
      Height          =   375
      Left            =   2430
      TabIndex        =   2
      Top             =   3960
      Width           =   1140
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   4770
      TabIndex        =   16
      Top             =   45
      Width           =   2580
   End
   Begin VB.DirListBox Dir1 
      Height          =   4290
      Left            =   4770
      TabIndex        =   15
      Top             =   405
      Width           =   2580
   End
   Begin VB.ComboBox cbxFormat 
      Height          =   300
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   270
      Width           =   3525
   End
   Begin VB.Timer tmrRecord 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5355
      Top             =   6480
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   990
      TabIndex        =   8
      Top             =   4770
      Width           =   6360
   End
   Begin VB.ListBox lstRecord 
      Height          =   2940
      Left            =   2385
      TabIndex        =   4
      Top             =   945
      Width           =   2295
   End
   Begin VB.ListBox lstTracks 
      Height          =   2940
      Left            =   45
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   945
      Width           =   2295
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新CD信息(&R)"
      Height          =   375
      Left            =   45
      TabIndex        =   18
      Top             =   3960
      Width           =   2265
   End
   Begin VB.Label lblCDID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1440
      TabIndex        =   9
      Top             =   45
      Width           =   105
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "保存质量(&S):"
      Height          =   195
      Index           =   5
      Left            =   45
      TabIndex        =   13
      Top             =   315
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1575
      TabIndex        =   12
      Top             =   7155
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "保存为(A):"
      Height          =   180
      Index           =   4
      Left            =   45
      TabIndex        =   7
      Top             =   4815
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "要保存的曲目(&H):"
      Height          =   195
      Index           =   3
      Left            =   2385
      TabIndex        =   3
      Top             =   720
      Width           =   2130
   End
   Begin VB.Label Label1 
      Caption         =   "不保存的曲目(&C):"
      Height          =   195
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   720
      Width           =   1905
   End
   Begin VB.Label lblNumTracks 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      TabIndex        =   11
      Top             =   45
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "CD ID:                              共       首曲目"
      Height          =   180
      Index           =   2
      Left            =   135
      TabIndex        =   10
      Top             =   45
      Width           =   4590
   End
End
Attribute VB_Name = "Form101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Dim lCurTrack As Long, lTrackLengths() As Long, lStart As Long, lFinish As Long, aFile As String, bGroups As Boolean
Private Function CString(aStr As String) As String
    CString = ""
    Dim k As Long
    k = InStr(aStr, Chr$(0))
    If k Then
        CString = Left$(aStr, k - 1)
    End If
End Function
Private Sub StartTrackRecording()
    Dim lRet As Long, lBits As Long, lSamples As Long, lChannels As Long
    lCurTrack = lstRecord.ItemData(lstRecord.ListIndex)
    lblStatus.Caption = "Track " & lCurTrack
    aFile = txtFile.Text & "-" & lCurTrack & ".wav"
    lStart = 0
    lFinish = 0
    For lRet = 1 To Val(lblNumTracks.Caption)
        If lRet = lCurTrack Then Exit For
        lStart = lStart + lTrackLengths(lRet)
    Next
    lFinish = lStart + lTrackLengths(lCurTrack)
    If bGroups Then
        Do
            If lstRecord.ListCount - 1 = lstRecord.ListIndex Then Exit Do
            lstRecord.ListIndex = lstRecord.ListIndex + 1
            If lstRecord.List(lstRecord.ListIndex) = "--- Group ---" Then Exit Do
            lCurTrack = lstRecord.ItemData(lstRecord.ListIndex)
            lFinish = lFinish + lTrackLengths(lCurTrack)
        Loop
    End If
    Select Case cbxFormat.List(cbxFormat.ListIndex)
        Case " 8.000kHz   8bit Mono      7k/sec": lSamples = 8000: lBits = 8: lChannels = 1
        Case " 8.000kHz   8bit Stereo   15k/sec": lSamples = 8000: lBits = 8: lChannels = 2
        Case " 8.000kHz  16bit Mono     15k/sec": lSamples = 8000: lBits = 16: lChannels = 1
        Case " 8.000kHz  16bit Stereo   31k/sec": lSamples = 8000: lBits = 16: lChannels = 2
        
        Case "11.025kHz   8bit Mono     10k/sec": lSamples = 11025: lBits = 8: lChannels = 1
        Case "11.025kHz   8bit Stereo   21k/sec": lSamples = 11025: lBits = 8: lChannels = 2
        Case "11.025kHz  16bit Mono     21k/sec": lSamples = 11025: lBits = 16: lChannels = 1
        Case "11.025kHz  16bit Stereo   43k/sec": lSamples = 11025: lBits = 16: lChannels = 2
        
        Case "22.050 Hz   8bit Mono     21k/sec": lSamples = 22050: lBits = 8: lChannels = 1
        Case "22.050 Hz   8bit Stereo   43k/sec": lSamples = 22050: lBits = 8: lChannels = 2
        Case "22.050 Hz  16bit Mono     43k/sec": lSamples = 22050: lBits = 16: lChannels = 1
        Case "22.050 Hz  16bit Stereo   86k/sec": lSamples = 22050: lBits = 16: lChannels = 2
        
        Case "44.100 Hz   8bit Mono     43k/sec": lSamples = 44100: lBits = 8: lChannels = 1
        Case "44.100 Hz   8bit Stereo   86k/sec": lSamples = 44100: lBits = 8: lChannels = 2
        Case "44.100 Hz  16bit Mono     86k/sec": lSamples = 44100: lBits = 16: lChannels = 1
        Case "44.100 Hz  16bit Stereo  172k/sec": lSamples = 44100: lBits = 16: lChannels = 2
    End Select
    Dim iBlockAlign As Integer, lBytesPerSec As Long
    iBlockAlign = lChannels * lBits / 8
    lBytesPerSec = lSamples * iBlockAlign
    If mciSendString("open new type waveaudio alias capture", vbNullString, 0, 0) Then MsgBox "Error opening waveaudio", vbCritical: cmdCancel_Click
    If lFinish Then
        If mciSendString("set capture samplespersec " & lSamples & " channels " & lChannels & " bitspersample " & lBits & " alignment " & iBlockAlign & " bytespersec " & lBytesPerSec, vbNullString, 0, 0) Then MsgBox "Error setting capture recording paramaters", vbCritical: mciSendString "close capture", vbNullString, 0, 0: cmdCancel_Click
    End If
    If lFinish Then
        If mciSendString("open cdaudio alias cd", vbNullString, 0, 0) Then
            MsgBox "Error opening cd!", vbCritical: cmdCancel_Click
        Else
            mciSendString "set cd time format milliseconds", vbNullString, 0, 0
            mciSendString "record capture overwrite", vbNullString, 0, 0
            If lStart Then
                lRet = mciSendString("play cd from " & lStart, vbNullString, 0, 0)
            Else
                lRet = mciSendString("play cd", vbNullString, 0, 0)
            End If
            If lRet Then MsgBox "Error playing cd!", vbCritical: cmdCancel_Click
        End If
    End If
    tmrRecord.Enabled = True
End Sub
Private Sub cmdCancel_Click()
    lFinish = 0
    lstRecord.ListIndex = lstRecord.ListCount - 1
End Sub
Private Sub cmdFrom_Click()
    On Local Error Resume Next
    If lstRecord.List(lstRecord.ListIndex) <> "--- Group ---" Then
        lstTracks.AddItem lstRecord.List(lstRecord.ListIndex)
        lstTracks.ItemData(lstTracks.NewIndex) = lstRecord.ItemData(lstRecord.ListIndex)
    End If
    lstRecord.RemoveItem lstRecord.ListIndex
End Sub
Private Sub cmdRefresh_Click()
    Dim aRet As String, lRet As Long, aTrack As String
    aRet = Space$(64)
    aTrack = Space$(2)
    lblCDID.Caption = ""
    lblNumTracks.Caption = ""
    lstTracks.Clear
    lstRecord.Clear
    If mciSendString("open cdaudio alias cd", vbNullString, 0, 0) = 0 Then
        mciSendString "info cd identity", aRet, 64, 0
        lblCDID.Caption = CString(aRet)
        txtFile.Text = App.Path & "\CD-" & lblCDID.Caption
        mciSendString "status cd number of tracks", aRet, 64, 0
        lblNumTracks.Caption = CString(aRet)
        mciSendString "set cd time format hms", vbNullString, 0, 0
        For lRet = 1 To Val(lblNumTracks.Caption)
            mciSendString "status cd length track " & lRet, aRet, 64, 0
            RSet aTrack = CStr(lRet)
            lstTracks.AddItem "Track " & aTrack & " - " & CString(aRet)
            lstTracks.ItemData(lstTracks.NewIndex) = lRet
        Next
       ReDim lTrackLengths(1 To Val(lblNumTracks.Caption)) As Long
        mciSendString "set cd time format milliseconds", vbNullString, 0, 0
        For lRet = 1 To Val(lblNumTracks.Caption)
            mciSendString "status cd length track " & lRet, aRet, 64, 0
           lTrackLengths(lRet) = CLng(CString(aRet))
        Next
        mciSendString "Close CD", vbNullString, 0, 0
        lstTracks.AddItem "--- 曲目信息 ---"
    End If
End Sub
Private Sub cmdStart_Click()
    If Len(txtFile.Text) = 0 Then MsgBox "You must enter a filename.", vbInformation: txtFile.SetFocus: Exit Sub
    If InStr(LCase$(txtFile.Text), ".wav") Then MsgBox "Don't include the .WAV extension.": txtFile.SetFocus: Exit Sub
    If lstRecord.ListCount = 0 Then MsgBox "You must select tracks to record.", vbInformation: lstTracks.SetFocus: Exit Sub
    Dim k As Long, bOutOfOrder As Boolean
    bGroups = False
    For k = 0 To lstRecord.ListCount - 1
        If lstRecord.List(k) = "--- Group ---" Then
            bGroups = True
        ElseIf k > 0 Then
            If lstRecord.ItemData(k - 1) <> lstRecord.ItemData(k) - 1 Then
                bOutOfOrder = True
            End If
        End If
    Next
    If bGroups And bOutOfOrder Then
        MsgBox "Tracks grouped together must be in sequence.", vbCritical
        Exit Sub
    End If
    lblStatus.Caption = ""
    lblStatus.Visible = True
    cmdCancel.Enabled = True
    cbxFormat.Enabled = False
    cmdStart.Enabled = False
    lstTracks.Enabled = False
    cmdRefresh.Enabled = False
    cmdTo.Enabled = False
    cmdFrom.Enabled = False
    lstRecord.Enabled = False
    txtFile.Enabled = False
    lstRecord.ListIndex = 0
    StartTrackRecording
End Sub
Private Sub cmdTo_Click()
    On Local Error Resume Next
    If lstTracks.List(lstTracks.ListIndex) = "--- Group ---" Then
        If lstRecord.ListCount = 0 Then
            MsgBox "You must first add some tracks.", vbInformation
            Exit Sub
        ElseIf lstRecord.List(lstRecord.ListCount - 1) = "--- Group ---" Then
            MsgBox "You must first add some more tracks.", vbInformation
            Exit Sub
        End If
    End If
    lstRecord.AddItem lstTracks.List(lstTracks.ListIndex)
    lstRecord.ItemData(lstRecord.NewIndex) = lstTracks.ItemData(lstTracks.ListIndex)
    If lstTracks.List(lstTracks.ListIndex) <> "--- Group ---" Then
        lstTracks.RemoveItem lstTracks.ListIndex
    End If
End Sub
Private Sub Dir1_Change()
Dir1.Path = Drive1
File1.Path = Dir1.Path
txtFile.Text = File1.Path + "\"
End Sub
Private Sub Drive1_Change()
Dir1.Path = Drive1
File1.Path = Dir1.Path
txtFile.Text = File1.Path + "\"
End Sub
Private Sub Form_Load()
Unload Form102
    cmdRefresh_Click
    cbxFormat.AddItem " 8.000kHz   8bit Mono      7k/sec"
    cbxFormat.AddItem " 8.000kHz   8bit Stereo   15k/sec"
    cbxFormat.AddItem " 8.000kHz  16bit Mono     15k/sec"
    cbxFormat.AddItem " 8.000kHz  16bit Stereo   31k/sec"
    
    cbxFormat.AddItem "11.025kHz   8bit Mono     10k/sec"
    cbxFormat.AddItem "11.025kHz   8bit Stereo   21k/sec"
    cbxFormat.AddItem "11.025kHz  16bit Mono     21k/sec"
    cbxFormat.AddItem "11.025kHz  16bit Stereo   43k/sec"
    cbxFormat.ListIndex = cbxFormat.NewIndex
    
    cbxFormat.AddItem "22.050 Hz   8bit Mono     21k/sec"
    cbxFormat.AddItem "22.050 Hz   8bit Stereo   43k/sec"
    cbxFormat.AddItem "22.050 Hz  16bit Mono     43k/sec"
    cbxFormat.AddItem "22.050 Hz  16bit Stereo   86k/sec"
    
    cbxFormat.AddItem "44.100 Hz   8bit Mono     43k/sec"
    cbxFormat.AddItem "44.100 Hz   8bit Stereo   86k/sec"
    cbxFormat.AddItem "44.100 Hz  16bit Mono     86k/sec"
    cbxFormat.AddItem "44.100 Hz  16bit Stereo  172k/sec"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If tmrRecord.Enabled Then
        cmdCancel_Click
        While tmrRecord.Enabled: DoEvents: Wend
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Form102.Show
End Sub
Private Sub lstRecord_DblClick()
    cmdFrom_Click
End Sub
Private Sub lstTracks_DblClick()
    cmdTo_Click
End Sub
Private Sub tmrRecord_Timer()
    Dim aRet As String, lRet As Long, lTrack As Long
    aRet = Space$(64)
    mciSendString "status cd position", aRet, 64, 0
    lRet = Val(CString(aRet))
    If lFinish Then
        mciSendString "status cd current track", aRet, 64, 0
        lTrack = Val(CString(aRet))
        lblStatus.Caption = "Track " & lTrack & "  -  " & Int((lRet - lStart) / (lFinish - lStart) * 100) & "%"
    End If
    If lRet >= lFinish Then
        tmrRecord.Enabled = False
        mciSendString "stop capture", vbNullString, 0, 0
        mciSendString "stop cd", vbNullString, 0, 0
        mciSendString "save capture " & aFile, vbNullString, 0, 0
        mciSendString "close capture", vbNullString, 0, 0
        mciSendString "close cd", vbNullString, 0, 0
        If lstRecord.ListIndex + 1 < lstRecord.ListCount Then
            lstRecord.ListIndex = lstRecord.ListIndex + 1
            StartTrackRecording
        Else
            If lFinish Then
                MsgBox "完成转换!", vbInformation
            Else
                MsgBox "已经停止!", vbCritical
            End If
            lblStatus.Visible = False
            cmdCancel.Enabled = False
            cbxFormat.Enabled = True
            cmdStart.Enabled = True
            lstTracks.Enabled = True
            cmdRefresh.Enabled = True
            cmdTo.Enabled = True
            cmdFrom.Enabled = True
            lstRecord.Enabled = True
            txtFile.Enabled = True
        End If
    End If
End Sub
Private Sub txtFile_GotFocus()
    txtFile.SelStart = 0
    txtFile.SelLength = Len(txtFile.Text)
End Sub
