VERSION 5.00
Begin VB.Form frmMinSize 
   Caption         =   "Form1"
   ClientHeight    =   4740
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   0
      Left            =   588
      TabIndex        =   0
      Top             =   336
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   1
      Left            =   2100
      TabIndex        =   1
      Top             =   336
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   2
      Left            =   3612
      TabIndex        =   2
      Top             =   336
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   3
      Left            =   5124
      TabIndex        =   3
      Top             =   336
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   4
      Left            =   6636
      TabIndex        =   4
      Top             =   336
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   5
      Left            =   588
      TabIndex        =   5
      Top             =   1344
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   6
      Left            =   2100
      TabIndex        =   6
      Top             =   1344
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   7
      Left            =   3612
      TabIndex        =   7
      Top             =   1344
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   8
      Left            =   5124
      TabIndex        =   8
      Top             =   1344
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   9
      Left            =   6636
      TabIndex        =   9
      Top             =   1344
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   10
      Left            =   588
      TabIndex        =   10
      Top             =   2352
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   11
      Left            =   2100
      TabIndex        =   11
      Top             =   2352
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   12
      Left            =   3612
      TabIndex        =   12
      Top             =   2352
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   13
      Left            =   5124
      TabIndex        =   13
      Top             =   2352
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   14
      Left            =   6636
      TabIndex        =   14
      Top             =   2352
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   15
      Left            =   588
      TabIndex        =   15
      Top             =   3360
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   16
      Left            =   2100
      TabIndex        =   16
      Top             =   3360
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   17
      Left            =   3612
      TabIndex        =   17
      Top             =   3360
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   18
      Left            =   5124
      TabIndex        =   18
      Top             =   3360
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
   Begin Project1.ctxTrackMouse ctxTrackMouse1 
      Height          =   768
      Index           =   19
      Left            =   6636
      TabIndex        =   19
      Top             =   3360
      Width           =   1272
      _ExtentX        =   2244
      _ExtentY        =   1355
   End
End
Attribute VB_Name = "frmMinSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z

#Const ImplNoIdeProtection = (MST_NO_IDE_PROTECTION <> 0)
#Const ImplSelfContained = True

'=========================================================================
' API
'=========================================================================

Private Const WM_GETMINMAXINFO              As Long = &H24
'--- for Modern Subclassing Thunk (MST)
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const SIGN_BIT                      As Long = &H80000000
'--- end MST

'--- for Modern Subclassing Thunk (MST)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcOrdinal As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#If Not ImplNoIdeProtection Then
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
#End If
#If ImplSelfContained Then
    Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
#End If
'--- end MST

Private Type POINTAPI
    X                   As Long
    Y                   As Long
End Type

Private Type MINMAXINFO
    ptReserved          As POINTAPI
    ptMaxSize           As POINTAPI
    ptMaxPosition       As POINTAPI
    ptMinTrackSize      As POINTAPI
    ptMaxTrackSize      As POINTAPI
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_pSubclass         As IUnknown
Private m_sngMinWidth       As Single
Private m_sngMinHeight      As Single

'=========================================================================
' Methods
'=========================================================================

Public Function SubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Attribute SubclassProc.VB_MemberFlags = "40"
    Dim uInfo           As MINMAXINFO
    
    #If hWnd And wParam And Handled Then '--- touch args
    #End If
    Select Case wMsg
    Case WM_GETMINMAXINFO
        Call CopyMemory(uInfo, ByVal lParam, LenB(uInfo))
        uInfo.ptMinTrackSize.X = m_sngMinWidth / Screen.TwipsPerPixelX
        uInfo.ptMinTrackSize.Y = m_sngMinHeight / Screen.TwipsPerPixelX
        Call CopyMemory(ByVal lParam, uInfo, LenB(uInfo))
        Handled = True
    End Select
End Function

'=========================================================================
' Controls events
'=========================================================================

Private Sub Form_Load()
    Set m_pSubclass = InitSubclassingThunk(hWnd, Me, InitAddressOfMethod(Me, 5).SubclassProc(0, 0, 0, 0, 0))
    m_sngMinWidth = Width
    m_sngMinHeight = Height
End Sub

Private Sub Form_Click()
    With New frmWaitCompletion
        .Show vbModal, Me
    End With
End Sub

Private Sub ctxTrackMouse1_MouseEnter(Index As Integer)
    Debug.Print "ctxTrackMouse1_MouseEnter, Index=" & Index & ", IsHot=" & ctxTrackMouse1(Index).IsHot, Timer
End Sub

Private Sub ctxTrackMouse1_MouseLeave(Index As Integer)
    Debug.Print "ctxTrackMouse1_MouseLeave, Index=" & Index & ", IsHot=" & ctxTrackMouse1(Index).IsHot, Timer
End Sub

'=========================================================================
' The Modern Subclassing Thunk (MST)
'=========================================================================

Private Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As frmMinSize
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
    If hThunk = 0 Then
        Exit Function
    End If
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitSubclassingThunk(ByVal hWnd As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgepwEBAAV1aLdCQUg8YIgz4AdC+L+oHHKBIQAIvCBQwREACri8IFSBEQAKuLwgVYERAAq4vCBYAREACruQkAAADzpYHCKBIQAFJqHP9SEFqL+IvCq7gBAAAAqzPAq4tEJAyri3QkFKWlM8Crg+8cagBX/3IM/3cM/1IYi0QkGIk4Xl+4XBIQAC1wEBAAwhAADx8Ai0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1GIsKUv9xDP9yDP9RHItUJASLClL/URQzwMIEAJBVi+yLVRj/QgT/QhiLQhg7QgR0b4tCEIXAdGiLCotBLIXAdDdS/9BaiUIIg/gBd1OFwHUJgX0MAwIAAHRGiwpS/1EwWoXAdTuLClJq8P9xJP9RKFqpAAAACHUoUjPAUFCNRCQEUI1EJARQ/3UU/3UQ/3UM/3UI/3IQ/1IUWVhahcl1E1KLCv91FP91EP91DP91CP9RIFr/ShhQUug4////WF3CGAAPHwA=" ' 9.6.2020 13:56:03
    Const THUNK_SIZE    As Long = 492
    Static hThunk       As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitSubclassingThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        Call DefSubclassProc(0, 0, 0, 0)                                            '--- load comctl32
        aParams(4) = GetProcByOrdinal(GetModuleHandle("comctl32"), 410)             '--- 410 = SetWindowSubclass ordinal
        aParams(5) = GetProcByOrdinal(GetModuleHandle("comctl32"), 412)             '--- 412 = RemoveWindowSubclass ordinal
        aParams(6) = GetProcByOrdinal(GetModuleHandle("comctl32"), 413)             '--- 413 = DefSubclassProc ordinal
        '--- for IDE protection
        Debug.Assert pvThunkIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitSubclassingThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, hWnd, 0, VarPtr(aParams(0)), VarPtr(InitSubclassingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function pvThunkIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvThunkIdeOwner = True
End Function

Private Function pvThunkAllocate(sText As String, Optional ByVal Size As Long) As Long
    Static Map(0 To &H3FF) As Long
    Dim baInput()       As Byte
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lPtr            As Long
    
    pvThunkAllocate = VirtualAlloc(0, IIf(Size > 0, Size, (Len(sText) \ 4) * 3), MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If pvThunkAllocate = 0 Then
        Exit Function
    End If
    '--- init decoding maps
    If Map(65) = 0 Then
        baInput = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
        For lIdx = 0 To UBound(baInput)
            lChar = baInput(lIdx)
            Map(&H0 + lChar) = lIdx * (2 ^ 2)
            Map(&H100 + lChar) = (lIdx And &H30) \ (2 ^ 4) Or (lIdx And &HF) * (2 ^ 12)
            Map(&H200 + lChar) = (lIdx And &H3) * (2 ^ 22) Or (lIdx And &H3C) * (2 ^ 6)
            Map(&H300 + lChar) = lIdx * (2 ^ 16)
        Next
    End If
    '--- base64 decode loop
    baInput = StrConv(Replace(Replace(sText, vbCr, vbNullString), vbLf, vbNullString), vbFromUnicode)
    lPtr = pvThunkAllocate
    For lIdx = 0 To UBound(baInput) - 3 Step 4
        lChar = Map(baInput(lIdx + 0)) Or Map(&H100 + baInput(lIdx + 1)) Or Map(&H200 + baInput(lIdx + 2)) Or Map(&H300 + baInput(lIdx + 3))
        Call CopyMemory(ByVal lPtr, lChar, 3)
        lPtr = (lPtr Xor SIGN_BIT) + 3 Xor SIGN_BIT
    Next
End Function

#If ImplSelfContained Then
Private Property Get pvThunkGlobalData(sKey As String) As Long
    Dim sBuffer     As String
    
    sBuffer = String$(50, 0)
    Call GetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, lValue)
End Property
#End If

