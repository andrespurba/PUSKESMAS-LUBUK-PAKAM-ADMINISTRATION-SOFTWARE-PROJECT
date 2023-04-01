Attribute VB_Name = "Module1"
Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Sub ReleaseCapture Lib "User32" ()
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2



