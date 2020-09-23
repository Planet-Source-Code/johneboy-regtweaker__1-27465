Attribute VB_Name = "noclose"
Option Explicit
Private Declare Function GetSystemMenu Lib "user32" _
        (ByVal hwnd As Long, ByVal bRevert As Long) _
        As Long
Private Declare Function RemoveMenu Lib "user32" _
        (ByVal hMenu As Long, ByVal nPosition As Long, _
        ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

Public Sub RemoveCloseMenu(frm As Form)
    Dim hSysMenu As Long
    ' Get the system menu for the form
    hSysMenu = GetSystemMenu(frm.hwnd, 0)
    ' Remove the close item
    Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
    ' and the seperator
    Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub


