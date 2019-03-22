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
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type CWPSTRUCT
    lParam              As Long
    wParam              As Long
    lMessage            As Long
    hWnd                As Long
End Type

Private m_pHookThunk            As IUnknown
Private m_pSubclassThunk        As IUnknown
Private WithEvents m_oList      As VB.ListBox
Attribute m_oList.VB_VarHelpID = -1

Private Property Get AddressOfHookProc() As frmTestHook
    Set AddressOfHookProc = InitAddressOfMethod(Me, 3)
End Property

Private Property Get AddressOfWndProc() As frmTestHook
    Set AddressOfWndProc = InitAddressOfMethod(Me, 4)
End Property

Private Sub Form_Click()
    Set m_pHookThunk = Nothing
End Sub

Private Sub Form_Load()
    Set m_pHookThunk = InitHookingThunk(WH_CALLWNDPROC, Me, AddressOfHookProc.HookProc(0, 0, 0))
    Set m_oList = Controls.Add("VB.ListBox", "lst")
    Debug.Print "m_oList.hWnd=" & m_oList.hWnd, Timer
End Sub

Public Function HookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim cwp             As CWPSTRUCT
    
    If nCode = HC_ACTION Then
        Call CopyMemory(cwp, ByVal lParam, Len(cwp))
        If cwp.lMessage = WM_CREATE Then
            Debug.Print "cwp.hWnd=" & cwp.hWnd, Timer
            Set m_pSubclassThunk = InitSubclassingThunk(cwp.hWnd, Me, AddressOfWndProc.ListWndProc(0, 0, 0, 0))
        End If
    End If
    HookProc = CallNextHookProc(m_pHookThunk, nCode, wParam, lParam)
End Function

Public Function ListWndProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If wMsg = WM_CREATE Then
        Debug.Print "wMsg=" & wMsg & " (WM_CREATE)", Timer
    Else
        Debug.Print "wMsg=" & wMsg, Timer
    End If
    ListWndProc = DefSubclassProc(hWnd, wParam, wParam, lParam)
End Function
