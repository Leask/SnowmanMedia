VERSION 5.00
Object = "{7D8AD1A3-781D-11D2-8E34-B68BBB0AA34F}#11.0#0"; "SmM_Tools.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1275
   Enabled         =   0   'False
   Icon            =   "媒体关联.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1110
   ScaleWidth      =   1275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin API控制大全.LyfTools LF1 
      Left            =   180
      Top             =   225
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetShortPathName Lib "kernel32" Alias _
        "GetShortPathNameA" (ByVal lpszLongPath As String, _
        ByVal lpszShortPath As String, ByVal cchBuffer As _
        Long) As Long
Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance = True Then End
LF1.CreateKey "HKEY_CLASSES_ROOT\Directory"
LF1.CreateKey "HKEY_CLASSES_ROOT\Directory\Shell"
LF1.CreateKey "HKEY_CLASSES_ROOT\Directory\Shell\SnowmanA"
LF1.CreateKey "HKEY_CLASSES_ROOT\Directory\Shell\SnowmanA\Command"
LF1.CreateKey "HKEY_CLASSES_ROOT\Directory\Shell\SnowmanB"
LF1.CreateKey "HKEY_CLASSES_ROOT\Directory\Shell\SnowmanB\Command"

LF1.SetStringValue "HKEY_CLASSES_ROOT\Directory\Shell\SnowmanA", "", "用 Snowman Media 播放(&P)   "
LF1.SetStringValue "HKEY_CLASSES_ROOT\Directory\Shell\SnowmanA\Command", "", App.Path + "\Snowman.exe %1                                        "
LF1.SetStringValue "HKEY_CLASSES_ROOT\Directory\Shell\SnowmanB", "", "加入 Snowman Media 播放列表(&A)      "
LF1.SetStringValue "HKEY_CLASSES_ROOT\Directory\Shell\SnowmanB\Command", "", App.Path + "\SmM_Add.exe %1                                        "


LF1.CreateKey "HKEY_CLASSES_ROOT\Snowman.Media"
LF1.CreateKey "HKEY_CLASSES_ROOT\Snowman.Media\DefaultIcon"
LF1.CreateKey "HKEY_CLASSES_ROOT\Snowman.Media\Shell\播放"
LF1.CreateKey "HKEY_CLASSES_ROOT\Snowman.Media\Shell\加入 Snowman Media 播放列表"
LF1.CreateKey "HKEY_CLASSES_ROOT\Snowman.Media\Shell\播放\command"
LF1.CreateKey "HKEY_CLASSES_ROOT\Snowman.Media\Shell\加入 Snowman Media 播放列表\command"
LF1.SetStringValue "HKEY_CLASSES_ROOT\Snowman.Media", "", "Snowman Media 媒体文件    "


LF1.SetStringValue "HKEY_CLASSES_ROOT\Snowman.Media\DefaultIcon", "", GetShortFileName(App.Path + "\SmM_Icons\002.ico")
If LF1.FileExists(LF1.GetSysPath + "\sndvol32.exe") = True And LF1.FileExists(LF1.GetSysPath + "\WindowsLogon.manifest") = True Then LF1.SetStringValue "HKEY_CLASSES_ROOT\Snowman.Media\DefaultIcon", "", GetShortFileName(App.Path + "\SmM_Icons\001.ico")


LF1.SetStringValue "HKEY_CLASSES_ROOT\Snowman.Media\Shell", "", "播放  "
LF1.SetStringValue "HKEY_CLASSES_ROOT\Snowman.Media\Shell\播放", "", "播放(&P)  "
LF1.SetStringValue "HKEY_CLASSES_ROOT\Snowman.Media\Shell\加入 Snowman Media 播放列表", "", "加入 Snowman Media 播放列表(&A)      "


LF1.SetStringValue "HKEY_CLASSES_ROOT\Snowman.Media\Shell\播放\command", "", App.Path + "\Snowman.exe %1                                        "
LF1.SetStringValue "HKEY_CLASSES_ROOT\Snowman.Media\Shell\加入 Snowman Media 播放列表\command", "", App.Path + "\SmM_Add.exe %1                                        "

If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "CVD") = True Then
LF1.CreateKey "HKEY_CLASSES_ROOT\.AudioCD"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.AudioCD", "", "AudioCD"
LF1.CreateKey "HKEY_CLASSES_ROOT\AudioCD"
LF1.SetStringValue "HKEY_CLASSES_ROOT\AudioCD", "", "音频 CD  "
LF1.CreateKey "HKEY_CLASSES_ROOT\AudioCD\Shell"
LF1.CreateKey "HKEY_CLASSES_ROOT\AudioCD\Shell\播放"
LF1.CreateKey "HKEY_CLASSES_ROOT\AudioCD\Shell\播放\command"
LF1.SetStringValue "HKEY_CLASSES_ROOT\AudioCD\Shell", "", "播放  "
LF1.SetStringValue "HKEY_CLASSES_ROOT\AudioCD\Shell\播放", "", "播放(&P)  "
LF1.SetStringValue "HKEY_CLASSES_ROOT\AudioCD\Shell\播放\command", "", GetShortFileName(App.Path + "\SmM_CD.exe")
LF1.CreateKey "HKEY_CLASSES_ROOT\.DVD"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.DVD", "", "DVD"
LF1.CreateKey "HKEY_CLASSES_ROOT\DVD"
LF1.SetStringValue "HKEY_CLASSES_ROOT\DVD", "", "视频 DVD"
LF1.CreateKey "HKEY_CLASSES_ROOT\Shell"
LF1.CreateKey "HKEY_CLASSES_ROOT\DVD\Shell\播放"
LF1.CreateKey "HKEY_CLASSES_ROOT\DVD\Shell\播放\command"
LF1.SetStringValue "HKEY_CLASSES_ROOT\DVD\Shell\播放", "", "播放(&P)  "
LF1.SetStringValue "HKEY_CLASSES_ROOT\DVD\Shell", "", "播放  "
LF1.SetStringValue "HKEY_CLASSES_ROOT\DVD\Shell\播放\command", "", GetShortFileName(App.Path + "\SmM_DVD.exe")
End If

If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "DATVOB") = True Then
LF1.CreateKey "HKEY_CLASSES_ROOT\.dat"
LF1.CreateKey "HKEY_CLASSES_ROOT\.vob"
LF1.CreateKey "HKEY_CLASSES_ROOT\.cda"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.dat", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.vob", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.cda", "", "Snowman.Media"
End If

LF1.CreateKey "HKEY_CLASSES_ROOT\.smm"
LF1.CreateKey "HKEY_CLASSES_ROOT\.smv"
LF1.CreateKey "HKEY_CLASSES_ROOT\.sma"
LF1.CreateKey "HKEY_CLASSES_ROOT\.sml"
LF1.CreateKey "HKEY_CLASSES_ROOT\.ilxz"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.smm", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.smv", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.sma", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.sml", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.ilxz", "", "Snowman.Media"

If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mwm") = True Then
LF1.CreateKey "HKEY_CLASSES_ROOT\.asf"
LF1.CreateKey "HKEY_CLASSES_ROOT\.asx"
LF1.CreateKey "HKEY_CLASSES_ROOT\.wax"
LF1.CreateKey "HKEY_CLASSES_ROOT\.wm"
LF1.CreateKey "HKEY_CLASSES_ROOT\.wma"
LF1.CreateKey "HKEY_CLASSES_ROOT\.wmd"
LF1.CreateKey "HKEY_CLASSES_ROOT\.wmv"
LF1.CreateKey "HKEY_CLASSES_ROOT\.wvx"
LF1.CreateKey "HKEY_CLASSES_ROOT\.wmp"
LF1.CreateKey "HKEY_CLASSES_ROOT\.wmx"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.asf", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.asx", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.wax", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.wm", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.wma", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.wmd", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.wmv", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.wvx", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.wmp", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.wmx", "", "Snowman.Media"
End If
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Qt") = True Then
LF1.CreateKey "HKEY_CLASSES_ROOT\.mov"
LF1.CreateKey "HKEY_CLASSES_ROOT\.qt"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.mov", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.qt", "", "Snowman.Media"
End If

If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mw") = True Then
LF1.CreateKey "HKEY_CLASSES_ROOT\.avi"
LF1.CreateKey "HKEY_CLASSES_ROOT\.wav"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.avi", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.wav", "", "Snowman.Media"
End If

If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Ivf") = True Then
LF1.CreateKey "HKEY_CLASSES_ROOT\.ivf"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.ivf", "", "Snowman.Media"
End If

If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Swf") = True Then
LF1.CreateKey "HKEY_CLASSES_ROOT\.swf"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.swf", "", "Snowman.Flash"
LF1.CreateKey "HKEY_CLASSES_ROOT\ShockwaveFlash.ShockwaveFlash"
LF1.CreateKey "HKEY_CLASSES_ROOT\ShockwaveFlash.ShockwaveFlash\DefaultIcon"
LF1.CreateKey "HKEY_CLASSES_ROOT\ShockwaveFlash.ShockwaveFlash\Shell\播放"
LF1.CreateKey "HKEY_CLASSES_ROOT\ShockwaveFlash.ShockwaveFlash\Shell\播放\command"
LF1.CreateKey "HKEY_CLASSES_ROOT\ShockwaveFlash.ShockwaveFlash\Shell\加入 Snowman Media 播放列表"
LF1.SetStringValue "HKEY_CLASSES_ROOT\ShockwaveFlash.ShockwaveFlash\Shell\加入 Snowman Media 播放列表", "", "加入 Snowman Media 播放列表(&A)      "
LF1.CreateKey "HKEY_CLASSES_ROOT\ShockwaveFlash.ShockwaveFlash\Shell\加入 Snowman Media 播放列表\command"
LF1.SetStringValue "HKEY_CLASSES_ROOT\ShockwaveFlash.ShockwaveFlash\Shell\加入 Snowman Media 播放列表\command", "", App.Path + "\SmM_Add.exe %1                                        "





LF1.SetStringValue "HKEY_CLASSES_ROOT\ShockwaveFlash.ShockwaveFlash", "", "Macromedia Flash 矢量影片文件      "
LF1.SetStringValue "HKEY_CLASSES_ROOT\ShockwaveFlash.ShockwaveFlash\DefaultIcon", "", GetShortFileName(App.Path + "\SmM_Icons\001.ico")
If LF1.FileExists(LF1.GetWinPath + "\sndvol32.exe") = True Then LF1.SetStringValue "HKEY_CLASSES_ROOT\Snowman.Media\DefaultIcon", "", GetShortFileName(App.Path + "\SmM_Icons\002.ico")
LF1.SetStringValue "HKEY_CLASSES_ROOT\ShockwaveFlash.ShockwaveFlash\Shell", "", "播放  "
LF1.SetStringValue "HKEY_CLASSES_ROOT\ShockwaveFlash.ShockwaveFlash\Shell\播放", "", "播放(&P)  "
LF1.SetStringValue "HKEY_CLASSES_ROOT\ShockwaveFlash.ShockwaveFlash\Shell\播放\command", "", App.Path + "\SmM_Flash.exe %1                                          "
End If

If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Aiff") = True Then
LF1.CreateKey "HKEY_CLASSES_ROOT\.aif"
LF1.CreateKey "HKEY_CLASSES_ROOT\.aifc"
LF1.CreateKey "HKEY_CLASSES_ROOT\.aiff"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.aif", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.aifc", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.aiff", "", "Snowman.Media"
End If

If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mpeg") = True Then
LF1.CreateKey "HKEY_CLASSES_ROOT\.mpeg"
LF1.CreateKey "HKEY_CLASSES_ROOT\.mpg"
LF1.CreateKey "HKEY_CLASSES_ROOT\.m1v"
LF1.CreateKey "HKEY_CLASSES_ROOT\.mp2"
LF1.CreateKey "HKEY_CLASSES_ROOT\.mpa"
LF1.CreateKey "HKEY_CLASSES_ROOT\.mpe"
LF1.CreateKey "HKEY_CLASSES_ROOT\.mp2v"
LF1.CreateKey "HKEY_CLASSES_ROOT\.mpv2"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.mpeg", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.mpg", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.m1v", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.mp2", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.mpa", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.mpe", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.mp2v", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.mpv2", "", "Snowman.Media"
End If
If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Rm") = True Then
LF1.CreateKey "HKEY_CLASSES_ROOT\.ra"
LF1.CreateKey "HKEY_CLASSES_ROOT\.rm"
LF1.CreateKey "HKEY_CLASSES_ROOT\.rmm"
LF1.CreateKey "HKEY_CLASSES_ROOT\.r1m"
LF1.CreateKey "HKEY_CLASSES_ROOT\.rom"
LF1.CreateKey "HKEY_CLASSES_ROOT\.mns"
LF1.CreateKey "HKEY_CLASSES_ROOT\.rp"
LF1.CreateKey "HKEY_CLASSES_ROOT\.rtx"
LF1.CreateKey "HKEY_CLASSES_ROOT\.rt"
LF1.CreateKey "HKEY_CLASSES_ROOT\.rmx"
LF1.CreateKey "HKEY_CLASSES_ROOT\.ram"
LF1.CreateKey "HKEY_CLASSES_ROOT\.rmj"
LF1.CreateKey "HKEY_CLASSES_ROOT\.rms"
LF1.CreateKey "HKEY_CLASSES_ROOT\.pls"
LF1.CreateKey "HKEY_CLASSES_ROOT\.xpl"
LF1.CreateKey "HKEY_CLASSES_ROOT\.smi"
LF1.CreateKey "HKEY_CLASSES_ROOT\.smil"
LF1.CreateKey "HKEY_CLASSES_ROOT\.mnd"
LF1.CreateKey "HKEY_CLASSES_ROOT\.rmvb"
LF1.CreateKey "HKEY_CLASSES_ROOT\.ssm"
LF1.CreateKey "HKEY_CLASSES_ROOT\.rv"
LF1.CreateKey "HKEY_CLASSES_ROOT\.sdp"
LF1.CreateKey "HKEY_CLASSES_ROOT\.r3t"
LF1.CreateKey "HKEY_CLASSES_ROOT\.acp"
LF1.CreateKey "HKEY_CLASSES_ROOT\.la1"
LF1.CreateKey "HKEY_CLASSES_ROOT\.lar"
LF1.CreateKey "HKEY_CLASSES_ROOT\.vpg"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.ra", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.rm", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.rmm", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.r1m", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.rom", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.mns", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.rp", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.rtx", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.rt", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.rmx", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.ram", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.rmj", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.rms", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.pls", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.xpl", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.smi", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.smil", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.mnd", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.rmvb", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.ssm", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.rv", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.sdp", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.r3t", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.acp", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.la1", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.lar", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.vpg", "", "Snowman.Media"
End If

If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Au") = True Then
LF1.CreateKey "HKEY_CLASSES_ROOT\.au"
LF1.CreateKey "HKEY_CLASSES_ROOT\.snd"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.au", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.snd", "", "Snowman.Media"
End If

If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Mp3") = True Then
LF1.CreateKey "HKEY_CLASSES_ROOT\.mp3"
LF1.CreateKey "HKEY_CLASSES_ROOT\.m3u"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.mp3", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.m3u", "", "Snowman.Media"
End If

If LF1.GetDWORDValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2O Networks\Snowman Media ilxz 4", "SetMt_Midi") = True Then
LF1.CreateKey "HKEY_CLASSES_ROOT\.mid"
LF1.CreateKey "HKEY_CLASSES_ROOT\.midi"
LF1.CreateKey "HKEY_CLASSES_ROOT\.rmi"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.mid", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.midi", "", "Snowman.Media"
LF1.SetStringValue "HKEY_CLASSES_ROOT\.rmi", "", "Snowman.Media"
End If

End
End Sub
Public Function GetShortFileName(ByVal FileName As String) As String
    
    Dim rc As Long
    Dim ShortPath As String
    Const PATH_LEN& = 164
    '获得文件的短文件名
    ShortPath = String$(PATH_LEN + 1, 0)
    rc = GetShortPathName(FileName, ShortPath, PATH_LEN)
    GetShortFileName = Left$(ShortPath, rc)
End Function

