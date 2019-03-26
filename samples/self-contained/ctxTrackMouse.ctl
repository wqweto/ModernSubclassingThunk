VERSION 5.00
Begin VB.UserControl ctxTrackMouse 
   BackColor       =   &H80000018&
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   2880
   ScaleWidth      =   3840
End
Attribute VB_Name = "ctxTrackMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z

#Const ImplNoVBIDESupport = (MST_NO_VBIDE_SUPPORT <> 0)

'=========================================================================
' Events
'=========================================================================

Event MouseEnter()
Event MouseLeave()

'=========================================================================
' API
'=========================================================================

Private Const WM_CANCELMODE                 As Long = &H1F
Private Const WM_MOUSELEAVE                 As Long = &H2A3
'--- for Modern Subclassing Thunk (MST)
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const CRYPT_STRING_BASE64           As Long = 1
'--- end MST

Private Declare Function TrackMouseEvent Lib "comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSESTRUCT) As Long
'--- for Modern Subclassing Thunk (MST)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryA" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, pcbBinary As Long, Optional ByVal pdwSkip As Long, Optional ByVal pdwFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcAddressByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcOrdinal As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#If Not ImplNoVBIDESupport Then
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If
'--- end MST

Private Enum TrackMouseEventFlags
    TME_HOVER = &H1
    TME_LEAVE = &H2
    TME_NONCLIENT = &H10
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSESTRUCT
    cbSize              As Long
    dwFlags             As TrackMouseEventFlags
    hwndTrack           As Long
    dwHoverTime         As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_pSubclass         As IUnknown
Private m_bIsHot            As Boolean

'=========================================================================
' Properties
'=========================================================================

Property Get IsHot() As Boolean
    IsHot = m_bIsHot
End Property

Property Let IsHot(ByVal bValue As Boolean)
    m_bIsHot = bValue
    BackColor = IIf(bValue, vbHighlight, vbInfoBackground)
End Property

'=========================================================================
' Methods
'=========================================================================

Private Sub pvSubclass()
    Set m_pSubclass = InitSubclassingThunk(hWnd, Me, InitAddressOfMethod(Me, 4).SubclassProc(0, 0, 0, 0))
End Sub

Private Sub pvUnsubclass()
    Set m_pSubclass = Nothing
End Sub

Public Function SubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Attribute SubclassProc.VB_MemberFlags = "40"
    Select Case wMsg
    Case WM_MOUSELEAVE, WM_CANCELMODE
        If IsHot Then
            IsHot = False
            RaiseEvent MouseLeave
        End If
    End Select
    SubclassProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
End Function

'=========================================================================
' Control events
'=========================================================================

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim uTrackMouse     As TRACKMOUSESTRUCT
    
    If Not IsHot Then
        IsHot = True
        With uTrackMouse
            .cbSize = Len(uTrackMouse)
            .hwndTrack = hWnd
            .dwFlags = TME_LEAVE Or TME_HOVER
        End With
        Call TrackMouseEvent(uTrackMouse)
        RaiseEvent MouseEnter
    End If
End Sub

Private Sub UserControl_InitProperties()
    pvSubclass
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    pvSubclass
End Sub

Private Sub UserControl_Terminate()
    pvUnsubclass
End Sub

'=========================================================================
' The Modern Subclassing Thunk (MST)
'=========================================================================

Private Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As ctxTrackMouse
    Const STR_THUNK     As String = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitSubclassingThunk(ByVal hWnd As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Const STR_THUNK     As String = "6AAAAABag+oFgepwECQBV1aLdCQUg8YIgz4AdC+L+oHH3BEkAYvCBQQRJAGri8IFQBEkAauLwgVQESQBq4vCBXgRJAGruQkAAADzpYHC3BEkAVJqFP9SEFqL+IvCq7gBAAAAq4tEJAyri3QkFKWlg+8UagBX/3IM/3cI/1IYi0QkGIk4Xl+4EBIkAS1wECQBwhAAkItEJAiDOAB1KoN4BAB1JIF4CMAAAAB1G4F4DAAAAEZ1EotUJAT/QgSLRCQMiRAzwMIMALgCQACAwgwAkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEdRiLClL/cQz/cgj/URyLVCQEiwpS/1EUM8DCBACQVYvsi1UYiwqLQSyFwHQnUv/QWoP4AXU3iwpS/1EwWoXAdSyLClJq8P9xJP9RKFqpAAAACHUZM8BQVP91FP91EP91DP91CP9yDP9SEFjrEYsK/3UU/3UQ/3UM/3UI/1EgXcIYAA==" ' 25.3.2019 14:02:14
    Const THUNK_SIZE    As Long = 416
    Static hThunk       As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    If hThunk = 0 Then
        hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        Call DefSubclassProc(0, 0, 0, 0)                                            '--- load comctl32
        aParams(4) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 410)      '--- 410 = SetWindowSubclass ordinal
        aParams(5) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 412)      '--- 412 = RemoveWindowSubclass ordinal
        aParams(6) = GetProcAddressByOrdinal(GetModuleHandle("comctl32"), 413)      '--- 413 = DefSubclassProc ordinal
        #If Not ImplNoVBIDESupport Then
            aParams(7) = hIdeOwner
            If aParams(7) <> 0 Then
                aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
                aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
                aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
            End If
        #End If
    End If
    lSize = CallWindowProc(hThunk, hWnd, 0, VarPtr(aParams(0)), VarPtr(InitSubclassingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

#If Not ImplNoVBIDESupport Then
Private Property Get hIdeOwner() As Long
    Dim lProcessId      As Long
    
    Do
        hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
        Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
    Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
End Property
#End If
