VERSION 5.00
Begin VB.Form FmSt 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Snowman Media  3.0"
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   Icon            =   "FmSt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Visible         =   0   'False
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   """"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3015
      TabIndex        =   0
      Top             =   1485
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "FmSt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Filename As String
Private Sub Form_Load()

On Error Resume Next

If App.PrevInstance = False Then
Form102.Show
End If
End Sub
