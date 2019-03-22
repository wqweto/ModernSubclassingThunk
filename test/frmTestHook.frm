VERSION 5.00
Begin VB.Form frmTestHook 
   Caption         =   "Form2"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form2"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmTestHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WM_CREATE                     As Long = &H1
'--- for windows hooks
Private Const WH_CALLWNDPROC                As Long = 4
Private Const HC_ACTION                     As Long = 0

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type CWPSTRUCT
    lParam              As Long
    wParam              As Long
    lMessage            As Long
    hWnd                As Long
End Type

Private m_pHook                 As IUnknown
Private m_pSubclass             As IUnknown
Private WithEvents m_oList      As VB.ListBox
Attribute m_oList.VB_VarHelpID = -1

Private Property Get pvAddressOfHookProc() As frmTestHook
    Set pvAddressOfHookProc = InitAddressOfMethod(Me, 3)
End Property

Private Property Get pvAddressOfSubclassProc() As frmTestHook
    Set pvAddressOfSubclassProc = InitAddressOfMethod(Me, 4)
End Property

Private Sub Form_Click()
    Set m_pHook = Nothing
End Sub

Private Sub Form_Load()
    Set m_pHook = InitHookingThunk(WH_CALLWNDPROC, Me, pvAddressOfHookProc.CallWndProcHookProc(0, 0, 0))
    Set m_oList = Controls.Add("VB.ListBox", "lst")
    Debug.Print "m_oList.hWnd=" & m_oList.hWnd, Timer
End Sub

Public Function CallWndProcHookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim cwp             As CWPSTRUCT
    
    If nCode = HC_ACTION Then
        Call CopyMemory(cwp, ByVal lParam, Len(cwp))
        If cwp.lMessage = WM_CREATE Then
            Debug.Print "cwp.hWnd=" & cwp.hWnd, Timer
            Set m_pSubclass = InitSubclassingThunk(cwp.hWnd, Me, pvAddressOfSubclassProc.ListSubclassProc(0, 0, 0, 0))
        End If
    End If
    CallWndProcHookProc = CallNextHookProc(m_pHook, nCode, wParam, lParam)
End Function

Public Function ListSubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If wMsg = WM_CREATE Then
        Debug.Print "hWnd=" & hWnd & ", wMsg=" & wMsg & " (WM_CREATE)", Timer
    Else
        Debug.Print "hWnd=" & hWnd & ", wMsg=" & wMsg, Timer
    End If
    ListSubclassProc = CallNextSubclassProc(m_pSubclass, hWnd, wParam, wParam, lParam)
End Function
