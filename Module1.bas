Attribute VB_Name = "Module1"
Declare Function SendMessage Lib "user32.dll" _
Alias "SendMessageA" (ByVal hWnd As Long, _
ByVal Msg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long

Public Const LVM_FIRST = &H1000
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2


