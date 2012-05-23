Attribute VB_Name = "Module1"
Option Explicit
Declare Function MCISendStringA& Lib "MMSYSTEM" (ByVal LPSTRCOMMAND$, ByVal LPSTRRETURNSTR As Any, ByVal WRETURNLEN%, ByVal HCALLBACK%)
