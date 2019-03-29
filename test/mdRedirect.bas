Attribute VB_Name = "mdRedirect"
Option Explicit

Public Function RedirectFrmTestSubclassWndProc(ByVal This As frmTestSubclass, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean, lResult As Long) As Long
    On Error GoTo EH
    lResult = This.frWndProc(hWnd, wMsg, wParam, lParam, Handled)
    Exit Function
EH:
    Debug.Print "Critical error: " & Err.Description
End Function

