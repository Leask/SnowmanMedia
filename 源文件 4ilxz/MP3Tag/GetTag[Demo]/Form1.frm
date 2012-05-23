VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form10 
   Caption         =   "GetMP3Tag[Deom]"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   6870
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "退出(&E)"
      Height          =   345
      Left            =   3240
      TabIndex        =   14
      Top             =   3330
      Width           =   1005
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保存(&S)"
      Enabled         =   0   'False
      Height          =   345
      Left            =   1680
      TabIndex        =   13
      Top             =   3330
      Width           =   1005
   End
   Begin VB.ComboBox CmbGenra 
      Height          =   300
      Left            =   2070
      TabIndex        =   6
      Top             =   2190
      Width           =   2145
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   2730
      Width           =   4035
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   2190
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Top             =   1590
      Width           =   4035
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   1020
      Width           =   4035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   4035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开(&O)"
      Height          =   345
      Left            =   180
      TabIndex        =   0
      Top             =   3330
      Width           =   1005
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   2430
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "关于..."
      Height          =   180
      Left            =   3630
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   60
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "注释："
      Height          =   180
      Index           =   5
      Left            =   210
      TabIndex        =   12
      Top             =   2550
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "风格："
      Height          =   180
      Index           =   4
      Left            =   2100
      TabIndex        =   11
      Top             =   1980
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "年代："
      Height          =   180
      Index           =   3
      Left            =   210
      TabIndex        =   10
      Top             =   1980
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "专辑名："
      Height          =   180
      Index           =   2
      Left            =   210
      TabIndex        =   9
      Top             =   1380
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "艺术家："
      Height          =   180
      Index           =   1
      Left            =   210
      TabIndex        =   8
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "歌曲名："
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   7
      Top             =   270
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   1
      X1              =   150
      X2              =   4250
      Y1              =   3165
      Y2              =   3165
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   4250
      Y1              =   3180
      Y2              =   3180
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFileName As String

Private Sub Command1_Click()

Dim strSongName As String
Dim strArtist As String
Dim strAlbum As String
Dim strYear As String
Dim strComment As String
Dim ByteGenra As Byte
Dim i As Byte

    With CommonDialog1
         .CancelError = True
         On Error GoTo ErrExit
         .Filter = "(MP3文件 *.MP3)|*.mp3"
         .DialogTitle = "打开MP3文件"
         .ShowOpen
         strFileName = .Filename
    End With
    
    With CmbGenra
        For i = 0 To 147
            .AddItem SetGenra(i)
        Next i
        .AddItem "", 148
        .ListIndex = 148
    End With
    
    GetTag1.GetMP3Tag strFileName, strSongName, strArtist, strAlbum, strYear, strComment, ByteGenra

    Text1.Text = strSongName
    Text2.Text = strArtist
    Text3.Text = strAlbum
    Text4.Text = strYear
    Text5.Text = strComment
    CmbGenra.ListIndex = GetGenraComboIdx(ByteGenra)
    
    Command2.Enabled = True

ErrExit:

End Sub

Private Function GetGenraComboIdx(ByVal nIndex As Byte) As Integer
    
    Dim i As Integer
    Dim GenraNameStr As String
    
    GenraNameStr = SetGenra(nIndex)
    With CmbGenra
        For i = 0 To .ListCount - 1
            If .List(i) = GenraNameStr Then
                Exit For
            End If
        Next i
        GetGenraComboIdx = IIf(i > .ListCount - 1, 148, i)
    End With
    
End Function

Private Function GetGenraCode(ByVal GenraName As String) As Byte
    
    Dim i As Byte
    
    For i = 0 To 147
        If GenraName = SetGenra(i) Then
            Exit For
        End If
    Next i
    GetGenraCode = IIf(i > 147, 255, i)
    
End Function

Private Function SetGenra(ByVal GenraComb As Byte) As String
    
    On Local Error Resume Next
    Select Case GenraComb
        Case Is = 0: SetGenra = "Blues"
        Case Is = 1: SetGenra = "Classic Rock"
        Case Is = 2: SetGenra = "Country"
        Case Is = 3: SetGenra = "Dance"
        Case Is = 4: SetGenra = "Disco"
        Case Is = 5: SetGenra = "Funk"
        Case Is = 6: SetGenra = "Grunge"
        Case Is = 7: SetGenra = "Hip Hop"
        Case Is = 8: SetGenra = "Jazz"
        Case Is = 9: SetGenra = "Metal"
        Case Is = 10: SetGenra = "New Age"
        Case Is = 11: SetGenra = "Oldies"
        Case Is = 12: SetGenra = "Other"
        Case Is = 13: SetGenra = "Pop"
        Case Is = 14: SetGenra = "Rhythm & Blues"
        Case Is = 15: SetGenra = "Rap"
        Case Is = 16: SetGenra = "Reggae"
        Case Is = 17: SetGenra = "Rock"
        Case Is = 18: SetGenra = "Techno"
        Case Is = 19: SetGenra = "Industrial"
        Case Is = 20: SetGenra = "Alternative"
        Case Is = 21: SetGenra = "Ska"
        Case Is = 22: SetGenra = "Death Metal"
        Case Is = 23: SetGenra = "Pranks"
        Case Is = 24: SetGenra = "Soundtrack"
        Case Is = 25: SetGenra = "Euro Techno"
        Case Is = 26: SetGenra = "Ambient"
        Case Is = 27: SetGenra = "Trip-Hop"
        Case Is = 28: SetGenra = "Vocal"
        Case Is = 29: SetGenra = "Jazz Funk"
        Case Is = 30: SetGenra = "Fusion"
        Case Is = 31: SetGenra = "Trance"
        Case Is = 32: SetGenra = "Classical"
        Case Is = 33: SetGenra = "Instrumental"
        Case Is = 34: SetGenra = "Acid"
        Case Is = 35: SetGenra = "House"
        Case Is = 36: SetGenra = "Game"
        Case Is = 37: SetGenra = "Sound Clip"
        Case Is = 38: SetGenra = "Gospel"
        Case Is = 39: SetGenra = "Noise"
        Case Is = 40: SetGenra = "Alternative Rock"
        Case Is = 41: SetGenra = "Bass"
        Case Is = 42: SetGenra = "Soul"
        Case Is = 43: SetGenra = "Punk"
        Case Is = 44: SetGenra = "Space"
        Case Is = 45: SetGenra = "Meditative"
        Case Is = 46: SetGenra = "Instrumental Pop"
        Case Is = 47: SetGenra = "Instrumental Rock"
        Case Is = 48: SetGenra = "Ethnic"
        Case Is = 49: SetGenra = "Gothic"
        Case Is = 50: SetGenra = "Darkwave"
        Case Is = 51: SetGenra = "Techno Industrial"
        Case Is = 52: SetGenra = "Electronic"
        Case Is = 53: SetGenra = "Pop Folk"
        Case Is = 54: SetGenra = "Eurodance"
        Case Is = 55: SetGenra = "Dream"
        Case Is = 56: SetGenra = "Southern Rock"
        Case Is = 57: SetGenra = "Comedy"
        Case Is = 58: SetGenra = "Cult"
        Case Is = 59: SetGenra = "Gangsta"
        Case Is = 60: SetGenra = "Top 40"
        Case Is = 61: SetGenra = "Christian Rap"
        Case Is = 62: SetGenra = "Pop Funk"
        Case Is = 63: SetGenra = "Jungle"
        Case Is = 64: SetGenra = "Native American"
        Case Is = 65: SetGenra = "Cabaret"
        Case Is = 66: SetGenra = "New Wave"
        Case Is = 67: SetGenra = "Psychadelic"
        Case Is = 68: SetGenra = "Rave"
        Case Is = 69: SetGenra = "Show Tunes"
        Case Is = 70: SetGenra = "Trailer"
        Case Is = 71: SetGenra = "Lo-Fi"
        Case Is = 72: SetGenra = "Tribal"
        Case Is = 73: SetGenra = "Acid Punk"
        Case Is = 74: SetGenra = "Acid Jazz"
        Case Is = 75: SetGenra = "Polka"
        Case Is = 76: SetGenra = "Retro"
        Case Is = 77: SetGenra = "Musical"
        Case Is = 78: SetGenra = "Rock & Roll"
        Case Is = 79: SetGenra = "Hard Rock"
        Case Is = 80: SetGenra = "Folk"
        Case Is = 81: SetGenra = "Folk/Rock"
        Case Is = 82: SetGenra = "National Folk"
        Case Is = 83: SetGenra = "Swing"
        Case Is = 84: SetGenra = "Fast Fusion"
        Case Is = 85: SetGenra = "Bebob"
        Case Is = 86: SetGenra = "Latin"
        Case Is = 87: SetGenra = "Revival"
        Case Is = 88: SetGenra = "Celtic"
        Case Is = 89: SetGenra = "Bluegrass"
        Case Is = 90: SetGenra = "Avantgarde"
        Case Is = 91: SetGenra = "Gothic Rock"
        Case Is = 92: SetGenra = "Progressive Rock"
        Case Is = 93: SetGenra = "Psychedelic Rock"
        Case Is = 94: SetGenra = "Symphonic Rock"
        Case Is = 95: SetGenra = "Slow Rock"
        Case Is = 96: SetGenra = "Big Band"
        Case Is = 97: SetGenra = "Chorus"
        Case Is = 98: SetGenra = "Easy Listening"
        Case Is = 99: SetGenra = "Acoustic"
        Case Is = 100: SetGenra = "Humour"
        Case Is = 101: SetGenra = "Speech"
        Case Is = 102: SetGenra = "Chanson"
        Case Is = 103: SetGenra = "Opera"
        Case Is = 104: SetGenra = "Chamber Music"
        Case Is = 105: SetGenra = "Sonata"
        Case Is = 106: SetGenra = "Symphony"
        Case Is = 107: SetGenra = "Booty Bass"
        Case Is = 108: SetGenra = "Primus"
        Case Is = 109: SetGenra = "Porn Groove"
        Case Is = 110: SetGenra = "Satire"
        Case Is = 111: SetGenra = "Slow Jam"
        Case Is = 112: SetGenra = "Club"
        Case Is = 113: SetGenra = "Tango"
        Case Is = 114: SetGenra = "Samba"
        Case Is = 115: SetGenra = "Folklore"
        Case Is = 116: SetGenra = "Ballad"
        Case Is = 117: SetGenra = "Power Ballad"
        Case Is = 118: SetGenra = "Rhytmic Soul"
        Case Is = 119: SetGenra = "Freestyle"
        Case Is = 120: SetGenra = "Duet"
        Case Is = 121: SetGenra = "Punk Rock"
        Case Is = 122: SetGenra = "Drum Solo"
        Case Is = 123: SetGenra = "A Capella"
        Case Is = 124: SetGenra = "Euro House"
        Case Is = 125: SetGenra = "Dance Hall"
        Case Is = 126: SetGenra = "Goa"
        Case Is = 127: SetGenra = "Drum & Bass"
        Case Is = 128: SetGenra = "Club House"
        Case Is = 129: SetGenra = "Hardcore"
        Case Is = 130: SetGenra = "Terror"
        Case Is = 131: SetGenra = "Indie"
        Case Is = 132: SetGenra = "Brit Pop"
        Case Is = 133: SetGenra = "Negerpunk"
        Case Is = 134: SetGenra = "Polsk Punk"
        Case Is = 135: SetGenra = "Beat"
        Case Is = 136: SetGenra = "Christian Gangsta Rap"
        Case Is = 137: SetGenra = "Heavy Metal"
        Case Is = 138: SetGenra = "Black Metal"
        Case Is = 139: SetGenra = "Crossover"
        Case Is = 140: SetGenra = "Contemporary Christian"
        Case Is = 141: SetGenra = "Christian Rock"
        Case Is = 142: SetGenra = "Merengue"
        Case Is = 143: SetGenra = "Salsa"
        Case Is = 144: SetGenra = "Trash Metal"
        Case Is = 145: SetGenra = "Anime"
        Case Is = 146: SetGenra = "JPop"
        Case Is = 147: SetGenra = "Synth Pop"
        Case Is = 255: SetGenra = "None"
        Case Else: SetGenra = ""
    End Select
    
End Function

Private Sub Command2_Click()
    
    Dim strSongName As String
    Dim strArtist As String
    Dim strAlbum As String
    Dim strYear As String
    Dim strComment As String
    Dim ByteGenra As Byte
    
    If GetTag1.WriteMP3Tag(strFileName, Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, GetGenraCode(CmbGenra.Text)) = True Then
       MsgBox "“" + strFileName + "”的 ID3 Tag 已经保存。", _
            vbInformation Or vbOKOnly, Me.Caption
    Else
       MsgBox "　　无法保存 ID3 Tag。" + vbCrLf + vbCrLf + _
            "　　将要写入数据的文件被锁定。" + _
            "请检查该文件的属性设置，或查看本控件的帮助。", vbCritical Or vbOKOnly, Me.Caption
    
    End If
    
End Sub

Private Sub Command4_Click()
    
    Unload Me
    End
    
End Sub

Private Sub Label2_Click()
    
    GetTag1.About
    
End Sub
