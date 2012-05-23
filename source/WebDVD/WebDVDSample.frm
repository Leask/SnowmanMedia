VERSION 5.00
Object = "{38EE5CE1-4B62-11D3-854F-00A0C9C898E7}#1.0#0"; "mswebdvd.dll"
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "SmM_Snowflake.ocx"
Begin VB.Form frmWebDVDSample 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "S.m.M. DVD Player"
   ClientHeight    =   6345
   ClientLeft      =   -19125
   ClientTop       =   360
   ClientWidth     =   7605
   Icon            =   "WebDVDSample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   7605
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C08044&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   6630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7620
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C08044&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   600
         Left            =   45
         TabIndex        =   3
         Top             =   5760
         Width           =   7530
         Begin VB.ListBox lstMenus 
            Appearance      =   0  'Flat
            Height          =   390
            ItemData        =   "WebDVDSample.frx":2CFA
            Left            =   5805
            List            =   "WebDVDSample.frx":2CFC
            TabIndex        =   4
            Top             =   90
            Width           =   1455
         End
         Begin VB.Image Image10 
            Appearance      =   0  'Flat
            Height          =   510
            Left            =   3690
            Picture         =   "WebDVDSample.frx":2CFE
            Stretch         =   -1  'True
            ToolTipText     =   "重新读取"
            Top             =   45
            Width           =   465
         End
         Begin VB.Image Image9 
            Appearance      =   0  'Flat
            Height          =   510
            Left            =   3060
            Picture         =   "WebDVDSample.frx":303F
            Stretch         =   -1  'True
            ToolTipText     =   "激活按钮"
            Top             =   45
            Width           =   555
         End
         Begin VB.Image Image8 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   2115
            Picture         =   "WebDVDSample.frx":3380
            Stretch         =   -1  'True
            ToolTipText     =   "下一章节"
            Top             =   90
            Width           =   375
         End
         Begin VB.Image Image7 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1710
            Picture         =   "WebDVDSample.frx":36C1
            Stretch         =   -1  'True
            ToolTipText     =   "弹出"
            Top             =   90
            Width           =   375
         End
         Begin VB.Image Image6 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   1305
            Picture         =   "WebDVDSample.frx":3A02
            Stretch         =   -1  'True
            ToolTipText     =   "停止"
            Top             =   90
            Width           =   375
         End
         Begin VB.Image Image5 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   945
            Picture         =   "WebDVDSample.frx":3D43
            Stretch         =   -1  'True
            ToolTipText     =   "暂停"
            Top             =   90
            Width           =   330
         End
         Begin VB.Image Image4 
            Appearance      =   0  'Flat
            Height          =   465
            Left            =   450
            Picture         =   "WebDVDSample.frx":4084
            Stretch         =   -1  'True
            ToolTipText     =   "开始"
            Top             =   45
            Width           =   420
         End
         Begin VB.Label lblTimeTrackerValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "00:00:00"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4545
            TabIndex        =   5
            Top             =   180
            Width           =   1080
         End
         Begin VB.Image Image2 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   90
            Picture         =   "WebDVDSample.frx":43C5
            Stretch         =   -1  'True
            ToolTipText     =   "上一章节"
            Top             =   90
            Width           =   330
         End
         Begin VB.Image Image3 
            Appearance      =   0  'Flat
            Height          =   480
            Left            =   3105
            Picture         =   "WebDVDSample.frx":4706
            Top             =   45
            Width           =   1050
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            Height          =   435
            Left            =   90
            Picture         =   "WebDVDSample.frx":501C
            Top             =   45
            Width           =   2400
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   4500
            Top             =   90
            Width           =   1140
         End
      End
      Begin MSWEBDVDLibCtl.MSWebDVD MSWebDVD1 
         Height          =   5685
         Left            =   45
         TabIndex        =   2
         Top             =   45
         Width           =   7530
         _cx             =   13282
         _cy             =   10028
         DisableAutoMouseProcessing=   0   'False
         BackColor       =   1048592
         EnableResetOnStop=   0   'False
         ColorKey        =   1048592
         WindowlessActivation=   0   'False
      End
   End
   Begin ACTIVESKINLibCtl.SkinForm Sf1 
      Height          =   480
      Left            =   7920
      OleObjectBlob   =   "WebDVDSample.frx":5D99
      TabIndex        =   0
      Top             =   5940
      Width           =   480
   End
End
Attribute VB_Name = "frmWebDVDSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'*******************************************************************************
'*       This is a part of the Microsoft DXSDK Code Samples.
'*       Copyright (C) 1999-2000 Microsoft Corporation.
'*       All rights reserved.
'*       This source code is only intended as a supplement to
'*       Microsoft Development Tools and/or SDK documentation.
'*       See these sources for detailed information regarding the
'*       Microsoft samples programs.
'*******************************************************************************
Option Explicit
Option Base 0
Option Compare Text



' **************************************************************************************************************************************
' * PRIVATE INTERFACE- FORM EVENT HANDLERS
' *
' *
            ' ******************************************************************************************************************************
            ' * procedure name: Form_Load
            ' * procedure description:  Occurs when a form is loaded.
            ' *
            ' ******************************************************************************************************************************
            Private Sub Form_Load()
                If App.PrevInstance Then End
        
       Sf1.SkinPath = App.Path + "\SmM_Skin"
    
            On Local Error GoTo ErrLine
            
            With lstMenus
               .AddItem "主菜单", 0
               .AddItem "标题菜单", 1
               .AddItem "音频菜单", 2
               .AddItem "附加菜单", 3
               .AddItem "章节菜单", 4
               .AddItem "副菜单", 5
            End With
            
            'set the root menu selected
            lstMenus.Selected(0) = True
            Exit Sub
            
ErrLine:
            Err.Clear
            Exit Sub
            End Sub
            
            
            ' ******************************************************************************************************************************
            ' * procedure name: Form_QueryUnload
            ' * procedure description:  Occurs before a form or application closes.
            ' *
            ' ******************************************************************************************************************************
            Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
            On Local Error GoTo ErrLine
            
            Select Case UnloadMode
                 Case 0 'vbFormControlMenu
                             Me.Move Screen.Width * 8, Screen.Height * 8
                             Me.Visible = False: Call MSWebDVD1.Stop
                 Case 1 'vbFormCode
                             Me.Move Screen.Width * 8, Screen.Height * 8
                             Me.Visible = False: Call MSWebDVD1.Stop
                 Case 2 'vbAppWindows
                             Me.Move Screen.Width * 8, Screen.Height * 8
                             Me.Visible = False: Call MSWebDVD1.Stop
                 Case 3 'vbAppTaskManager
                             Me.Move Screen.Width * 8, Screen.Height * 8
                             Me.Visible = False: Call MSWebDVD1.Stop
                 Case 4 'vbFormMDIForm
                             Exit Sub
                 Case 5 'vbFormOwner
                             Exit Sub
            End Select
            Exit Sub
            
ErrLine:
            Err.Clear
            Exit Sub
            End Sub



' **************************************************************************************************************************************
' * PRIVATE INTERFACE- CONTROL EVENT HANDLERS
' *
' *
            ' ******************************************************************************************************************************
            ' * procedure name: cmdPlay_Click
            ' * procedure description:  Occurs when the user clicks the "Play" command button
            ' *
            ' ******************************************************************************************************************************
            
            
            ' ******************************************************************************************************************************
            ' * procedure name: cmdStop_Click
            ' * procedure description:  Occurs when the user clicks the "Stop" command button
            ' *
            ' ******************************************************************************************************************************
            
            
            ' ******************************************************************************************************************************
            ' * procedure name: cmdPause_Click
            ' * procedure description:  Occurs when the user clicks the "Pause" command button
            ' *
            ' ******************************************************************************************************************************
            
            
            ' ******************************************************************************************************************************
            ' * procedure name: cmdEject_Click
            ' * procedure description:  Occurs when the user clicks the "Eject" command button
            ' *
            ' ******************************************************************************************************************************
            
            
            ' ******************************************************************************************************************************
            ' * procedure name: cmdActivateButton_Click
            ' * procedure description:  Occurs when the user clicks the "ActivateButton" command button
            ' *
            ' ******************************************************************************************************************************
            Private Sub cmdActivateButton_Click()
            On Local Error GoTo ErrLine

            'activates the currently selected button (selected button is highlighted)
            Call MSWebDVD1.ActivateButton
            Exit Sub
            
ErrLine:
            Call MsgBox(Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): Err.Clear
            Exit Sub
            End Sub
            
            
            ' ******************************************************************************************************************************
            ' * procedure name: cmdPlayNextChapter_Click
            ' * procedure description:  Occurs when the user clicks the "PlayNextChapter" command button
            ' *
            ' ******************************************************************************************************************************
            Private Sub cmdPlayNextChapter_Click()
            On Local Error GoTo ErrLine

            'takes playback to next chapter within current title
            Call MSWebDVD1.PlayNextChapter
            Exit Sub
            
ErrLine:
            Call MsgBox(Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): Err.Clear
            Exit Sub
            End Sub
            
            
            ' ******************************************************************************************************************************
            ' * procedure name: cmdPlayPrevChapter_Click
            ' * procedure description:  Occurs when the user clicks the "PlayPrevChapter" command button
            ' *
            ' ******************************************************************************************************************************
            Private Sub cmdPlayPrevChapter_Click()
            On Local Error GoTo ErrLine

            'takes playback to previous chapter within current title
            Call MSWebDVD1.PlayPrevChapter
            Exit Sub
            
ErrLine:
            Call MsgBox(Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): Err.Clear
            Exit Sub
            End Sub
            
            
            ' ******************************************************************************************************************************
            ' * procedure name: cmdShowMenu_Click
            ' * procedure description:  Occurs when the user clicks the "ShowMenu" command button
            ' *
            ' ******************************************************************************************************************************
            
            
            ' ******************************************************************************************************************************
            ' * procedure name: cmdResume_Click
            ' * procedure description:  Occurs when the user clicks the "Resume" command button
            ' *
            ' ******************************************************************************************************************************

            

Private Sub lblChoices_Click()

End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Height < 6855 Then Me.Height = 6800
If Me.Width < 7725 Then Me.Width = 7715
Frame1.Width = Me.Width
Frame1.Height = Me.Height
Frame1.Visible = False
Frame2.Top = Me.Height - 1040
Frame2.Width = Me.Width

Image3.Left = Me.Width - 4715
Image9.Left = Me.Width - 4665
Image10.Left = Me.Width - 4035
lblTimeTrackerValue.Left = Me.Width - 3180
Shape1.Left = Me.Width - 3225
lstMenus.Left = Me.Width - 1920
MSWebDVD1.Width = Me.Width - 195
MSWebDVD1.Height = Me.Height - 1170
Frame1.Visible = True
End Sub

Private Sub Image10_Click()
            On Local Error GoTo ErrLine

                ' Resume playback
                Call MSWebDVD1.Resume
            Exit Sub
            
ErrLine:
            Call MsgBox(Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): Err.Clear
            Exit Sub

End Sub

Private Sub Image2_Click()
            On Local Error GoTo ErrLine

            'takes playback to previous chapter within current title
            Call MSWebDVD1.PlayPrevChapter
            Exit Sub
            
ErrLine:
            Call MsgBox(Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): Err.Clear
            Exit Sub

End Sub

Private Sub Image4_Click()
            On Local Error GoTo ErrLine
            
            'Start playback
            Call MSWebDVD1.Play
            Exit Sub
            
ErrLine:
            Call MsgBox(Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): Err.Clear
            Exit Sub

End Sub

Private Sub Image5_Click()
            On Local Error GoTo ErrLine

            'pause playback
            Call MSWebDVD1.Pause
            Exit Sub
            
ErrLine:
            Call MsgBox(Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): Err.Clear
            Exit Sub

End Sub

Private Sub Image6_Click()
            On Local Error GoTo ErrLine
            
            'stop playback
            Call MSWebDVD1.Stop
            Exit Sub
            
ErrLine:
            Call MsgBox(Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): Err.Clear
            Exit Sub

End Sub

Private Sub Image7_Click()
            On Local Error GoTo ErrLine
            
            'Eject disc from the drive
            Call MSWebDVD1.Eject
            Exit Sub
            
ErrLine:
            Call MsgBox(Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): Err.Clear
            Exit Sub

End Sub

Private Sub Image8_Click()
            On Local Error GoTo ErrLine

            'takes playback to next chapter within current title
            Call MSWebDVD1.PlayNextChapter
            Exit Sub
            
ErrLine:
            Call MsgBox(Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): Err.Clear
            Exit Sub

End Sub

Private Sub Image9_Click()
            On Local Error GoTo ErrLine

            'activates the currently selected button (selected button is highlighted)
            Call MSWebDVD1.ActivateButton
            Exit Sub
            
ErrLine:
            Call MsgBox(Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): Err.Clear
            Exit Sub

End Sub

Private Sub lstMenus_Click()
            On Local Error GoTo ErrLine
            
                Select Case lstMenus.ListIndex
                    Case 0: Call MSWebDVD1.ShowMenu(3)  'Root
                    Case 1: Call MSWebDVD1.ShowMenu(2)  'Title
                    Case 2: Call MSWebDVD1.ShowMenu(5)  'Audio
                    Case 3: Call MSWebDVD1.ShowMenu(6)  'Angle
                    Case 4: Call MSWebDVD1.ShowMenu(7)  'Chapter
                    Case 5: Call MSWebDVD1.ShowMenu(4)  'Subpicture
                End Select
            Exit Sub
            
ErrLine:
            Call MsgBox(Err.Description, vbOKOnly + vbExclamation + vbApplicationModal, App.Title): Err.Clear
            Exit Sub

End Sub

            ' ******************************************************************************************************************************
            ' * procedure name: MSWebDVD1_DVDNotify
            ' * procedure description:  DVD notification event- occurs when a notification arrives from the dvd control
            ' *
            ' ******************************************************************************************************************************
            Private Sub MSWebDVD1_DVDNotify(ByVal lEventCode As Long, ByVal lParam1 As Variant, ByVal lParam2 As Variant)
            On Local Error GoTo ErrLine
            
            If 282 = lEventCode Then '282 is the event code for the time event
               'pass in param1 to get you the current time-convert to hh:mm:ss:ff format with DVDTimeCode2BSTR API
               If lblTimeTrackerValue.Caption <> CStr(MSWebDVD1.DVDTimeCode2bstr(lParam1)) Then _
                  lblTimeTrackerValue.Caption = CStr(MSWebDVD1.DVDTimeCode2bstr(lParam1))
            End If
            Exit Sub
            
ErrLine:
            Err.Clear
            Exit Sub
            End Sub
