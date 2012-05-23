VERSION 5.00
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About H L Sm.M. ilxz"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   6180
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picScroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   3165
      Left            =   0
      ScaleHeight     =   211
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   412
      TabIndex        =   0
      Top             =   1170
      Width           =   6180
   End
   Begin ACTIVESKINLibCtl.SkinForm SkinForm1 
      Height          =   480
      Left            =   1890
      OleObjectBlob   =   "About.frx":1582
      TabIndex        =   1
      Top             =   5715
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   4875
      Left            =   0
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   412
      TabIndex        =   2
      Top             =   0
      Width           =   6180
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTickCount Lib "Kernel32" () As Long



Const DT_BOTTOM As Long = &H8
Const DT_CALCRECT As Long = &H400
Const DT_CENTER As Long = &H1
Const DT_EXPANDTABS As Long = &H40
Const DT_EXTERNALLEADING As Long = &H200
Const DT_LEFT As Long = &H0
Const DT_NOCLIP As Long = &H100
Const DT_NOPREFIX As Long = &H800
Const DT_RIGHT As Long = &H2
Const DT_SINGLELINE As Long = &H20
Const DT_TABSTOP As Long = &H80
Const DT_TOP As Long = &H0
Const DT_VCENTER As Long = &H4
Const DT_WORDBREAK As Long = &H10

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'the actual text to scroll. This could also be loaded in from a text file
Dim ScrollText As String
                             

Dim EndingFlag As Boolean
Private Sub RunMain()
ScrollText = "H2ont Leask                " & vbCrLf & _
             "         Snowman Media     " & vbCrLf & _
             "           " & Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "VolID") & vbCrLf & vbCrLf & _
            vbCrLf & vbCrLf & vbCrLf & vbCrLf & "策划   黄思夏" & vbCrLf & vbCrLf & "设计   黄思夏" & vbCrLf & vbCrLf & "制作   黄思夏" & vbCrLf & vbCrLf & "编程   黄思夏" & vbCrLf & vbCrLf & "出品   黄思夏" & vbCrLf & vbCrLf & vbCrLf & "功能测试   梁  睿  朱素瑶  岳  校  苏小宁" & vbCrLf & vbCrLf & "平台测试   梁  睿  冯  婷" & vbCrLf & vbCrLf & "标志设计   黄思夏" & vbCrLf & vbCrLf & "图标设计   梁  睿" & vbCrLf & vbCrLf & "启动画面   梁  睿" & vbCrLf & vbCrLf & "启动音乐   梁  睿" & vbCrLf & vbCrLf & "界面   黄思夏  ( 梁  睿 提供 Jurassic Skinflake )" & vbCrLf & vbCrLf & "主题曲   《 远  行 》   作者   邓亦明" & vbCrLf & vbCrLf & "网站设计   梁  睿" & vbCrLf & vbCrLf & "文字处理   朱素瑶" & vbCrLf & vbCrLf & "Flash 动画   秋山月工作室 岳  校" & vbCrLf & vbCrLf & vbCrLf & _
             "控件提供   微软(Microsoft)中国有限公司  " & vbCrLf & vbCrLf & "           Macromedia 国际有限公司      " & vbCrLf & vbCrLf & "           RealNetwork 国际有限公司     " & vbCrLf & vbCrLf & "           Gamesman 公司 / DIB 技术公司 " & vbCrLf & vbCrLf & "           Dj's Computer Labs 工作室    " & vbCrLf & vbCrLf & "           武汉华工力学硕士研究生 刘玉锋" & vbCrLf & vbCrLf & vbCrLf & "素材提供   北京源江科技开发有限公司      " & vbCrLf & vbCrLf & "           北京恒星科贸有限责任公司      " & vbCrLf & vbCrLf & "           北京百联美达美数码科技有限公司" & vbCrLf & vbCrLf & vbCrLf & _
             "特别鸣谢   梁  睿  朱素瑶  苏小宁  徐  艳  冯  婷" & vbCrLf & vbCrLf & _
             "           秋山月工作室 岳  效                   " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                "H2ont  Running Splendors!                     " & vbCrLf & _
                             "           Copyright (C) 2000-2001 H2ont Leask" & vbCrLf & _
                             "           http://www.h2ont.com               " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                             "Professional by H2ont Leask" & vbCrLf & vbCrLf & _
                             "         " & Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "VolDay") & "         " & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf

Dim LastFrameTime As Long
Const IntervalTime As Long = 40
Dim rt As Long
Dim DrawingRect As RECT
Dim UpperX As Long, UpperY As Long 'Upper left point of drawing rect
Dim RectHeight As Long

'show the form
frmAbout.Refresh

'Get the size of the drawing rectangle by suppying the DT_CALCRECT constant
rt = DrawText(picScroll.hdc, ScrollText, -1, DrawingRect, DT_CALCRECT)

If rt = 0 Then 'err
    MsgBox "Error scrolling text", vbExclamation
    EndingFlag = True
Else
    DrawingRect.Top = picScroll.ScaleHeight
    DrawingRect.Left = 0
    DrawingRect.Right = picScroll.ScaleWidth
    'Store the height of The rect
    RectHeight = DrawingRect.Bottom
    DrawingRect.Bottom = DrawingRect.Bottom + picScroll.ScaleHeight
End If


Do While Not EndingFlag
    
    If GetTickCount() - LastFrameTime > IntervalTime Then
                    
        picScroll.Cls
        
        DrawText picScroll.hdc, ScrollText, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
        
        'update the coordinates of the rectangle
        DrawingRect.Top = DrawingRect.Top - 1
        DrawingRect.Bottom = DrawingRect.Bottom - 1
        
        'control the scolling and reset if it goes out of bounds
        If DrawingRect.Top < -(RectHeight) Then 'time to reset
            DrawingRect.Top = picScroll.ScaleHeight
            DrawingRect.Bottom = RectHeight + picScroll.ScaleHeight
        End If
        
        picScroll.Refresh
        
        LastFrameTime = GetTickCount()
        
    End If
    
    DoEvents
Loop

Unload Me
Set frmAbout = Nothing

End Sub


Private Sub Form_Activate()
Dim r As Integer
If Form102.FileExists(Form102.Label3.Caption + "\SmM_MC\YUANXING.WAV") = True Then
Form102.LyfTools1.PlayWav Form102.Label3.Caption + "\SmM_MC\YUANXING.wav", True
End If
 SkinForm1.SkinPath = Form102.LyfTools1.GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\H2ont_Leask\Snowman Media ilxz 3.5", "Skin_Path")
If Form102.FileExists(Form102.Label3.Caption + "\SmM_PT\SmM_ALA.gif") = True Then Picture2.Picture = LoadPicture(Form102.Label3.Caption + "\SmM_PT\SmM_ALA.gif")
If Form102.FileExists(Form102.Label3.Caption + "\SmM_PT\SmM_ALB.gif") = True Then Me.picScroll.Picture = LoadPicture(Form102.Label3.Caption + "\SmM_PT\SmM_ALB.gif")
RunMain
End Sub

Private Sub Form_Load()
Form102.LyfTools1.MakeTop Me, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form102.LyfTools1.PlayWav "", False
    EndingFlag = True
     Set frmAbout = Nothing
End Sub
