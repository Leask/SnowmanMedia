Attribute VB_Name = "Module100000"
Option Explicit
Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp&, ByVal wPixTypes&) As Long
Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp&) As Long


