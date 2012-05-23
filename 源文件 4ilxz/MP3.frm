VERSION 5.00
Object = "{EFD2F1D5-45DF-11D5-AA97-92FD4D1A316A}#1.0#0"; "SmM_InfoEdit.ocx"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "修改标签"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "MP3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Default         =   -1  'True
      Height          =   690
      Left            =   3150
      TabIndex        =   18
      Top             =   4500
      Width           =   1140
   End
   Begin SMQ_GetTag.GetTag GetTag1 
      Left            =   2205
      Top             =   3150
      _ExtentX        =   503
      _ExtentY        =   609
   End
   Begin VB.ComboBox CmbGenra 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   2880
      TabIndex        =   4
      Top             =   1125
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1035
      TabIndex        =   5
      Top             =   2115
      Width           =   4020
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1035
      TabIndex        =   3
      Top             =   1170
      Width           =   1005
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1035
      TabIndex        =   2
      Top             =   810
      Width           =   4020
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1035
      TabIndex        =   1
      Top             =   495
      Width           =   4020
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1035
      TabIndex        =   0
      Top             =   180
      Width           =   4020
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "流派 :"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   4
      Left            =   2205
      TabIndex        =   17
      Top             =   1170
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " 未知比特率"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   9
      Left            =   2835
      TabIndex        =   14
      Top             =   1485
      Width           =   2220
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "无信息"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   11
      Left            =   1035
      TabIndex        =   16
      Top             =   1800
      Width           =   4020
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      X1              =   1035
      X2              =   5040
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "00:00:00"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   8
      Left            =   1035
      TabIndex        =   13
      ToolTipText     =   "00:00:00"
      Top             =   1485
      Width           =   1005
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "时间   :"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   7
      Left            =   180
      TabIndex        =   12
      Top             =   1485
      Width           =   1230
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      X1              =   1035
      X2              =   2025
      Y1              =   1395
      Y2              =   1395
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   1035
      X2              =   5040
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   1035
      X2              =   5040
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   1035
      X2              =   5040
      Y1              =   405
      Y2              =   405
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "版权   :"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   6
      Left            =   180
      TabIndex        =   11
      Top             =   1800
      Width           =   1140
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "描述   :"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   10
      Top             =   2115
      Width           =   1185
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "年代   :"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   9
      Top             =   1170
      Width           =   1185
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "唱片集 :"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   8
      Top             =   810
      Width           =   1140
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "艺术家 :"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   495
      Width           =   1185
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "标题   :"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   180
      Width           =   1185
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "质量 :"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   10
      Left            =   2205
      TabIndex        =   15
      Top             =   1485
      Width           =   825
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      Height          =   2445
      Left            =   45
      Top             =   45
      Width           =   5190
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFileName As String
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
        Case Is = 14: SetGenra = "Rhythm and Blues"
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
        Case Is = 127: SetGenra = "Drum and Bass"
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



Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim strSongName As String
Dim strArtist As String
Dim strAlbum As String
Dim strYear As String
Dim strComment As String
Dim ByteGenra As Byte
Dim i As Byte


Form1.Ly.MakeTop Form1, False

         strFileName = Form1.LF1.List(Form1.LF2.ListIndex)

Form1.Mp(1).Filename = strFileName
Text1.Text = Form1.SotPath(strFileName)
If Len(Form1.Mp(1).GetMediaInfoString(mpClipTitle)) > 0 Then Text1.Text = Form1.Mp(1).GetMediaInfoString(mpClipTitle)
If Len(Form1.Mp(1).GetMediaInfoString(mpClipAuthor)) > 0 Then Text2.Text = Form1.Mp(1).GetMediaInfoString(mpClipAuthor)
If Len(Form1.Mp(1).GetMediaInfoString(mpClipCopyright)) > 0 Then Label1(11).Caption = Form1.Mp(1).GetMediaInfoString(mpClipCopyright)
If Len(Form1.Mp(1).GetMediaInfoString(mpClipDescription)) > 0 Then Text5.Text = Form1.Mp(1).GetMediaInfoString(mpClipDescription)
Label1(8).Caption = Form1.Gtime(Form1.Mp(1).Duration)
If Form1.Mp(1).Bandwidth > 0 Then Label1(9).Caption = Str(Int(Form1.Mp(1).Bandwidth / 1000)) + " 千字节每秒"
If Form1.Mp(1).ImageSourceWidth > 0 Then
    Me.Caption = "[" + Str(Form1.Mp(1).ImageSourceWidth) + " ×" + Str(Form1.Mp(1).ImageSourceHeight) + " 视频 ] " + strFileName
Else
     Me.Caption = "[ 音频 ] " + strFileName
End If

If UCase(Right(strFileName, 4)) = ".MP3" Then
        For i = 0 To 147
            CmbGenra.AddItem SetGenra(i)
        Next i
        CmbGenra.AddItem "", 148
        CmbGenra.ListIndex = 148
    
   GetTag1.GetMP3Tag strFileName, strSongName, strArtist, strAlbum, strYear, strComment, ByteGenra

    Text1.Text = strSongName
    Text2.Text = strArtist
    Text3.Text = strAlbum
    Text4.Text = strYear
    Text5.Text = strComment
    CmbGenra.ListIndex = GetGenraComboIdx(ByteGenra)

Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
CmbGenra.Enabled = True
Else
If Text2.Text = "" Then Text2.Text = "未知艺术家"
If Text3.Text = "" Then Text3.Text = "未知唱片集"
If Text4.Text = "" Then Text4.Text = "无信息"
If Text5.Text = "" Then Text5.Text = "无信息"

End If

ErrExit:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Jd As Long
    Dim strSongName As String
    Dim strArtist As String
    Dim strAlbum As String
    Dim strYear As String
    Dim strComment As String
    Dim ByteGenra As Byte
    Dim Id As Integer
    Id = Form1.LF2.ListIndex
     Form1.Mp(1).Filename = "ilxz"
 If UCase(Right(strFileName, 4)) = ".MP3" Then

    If Form1.Mp(0).Filename = strFileName Then
       Jd = Form1.Mp(0).CurrentPosition
    Form1.Mp(0).Filename = "ilxz"
     GetTag1.WriteMP3Tag strFileName, Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, GetGenraCode(CmbGenra.Text)
      Form1.ReNameB Id
      Form1.Mp(0).Filename = strFileName
           Form1.Mp(0).CurrentPosition = Jd
  Else
        GetTag1.WriteMP3Tag strFileName, Text1.Text, Text2.Text, Text3.Text, Text4.Text, Text5.Text, GetGenraCode(CmbGenra.Text)
      Form1.ReNameB Id

  End If
End If
If Form1.a00308.Checked = True Then Form1.Ly.MakeTop Form1, True
Set Form3 = Nothing

End Sub

