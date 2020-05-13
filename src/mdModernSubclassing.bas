Attribute VB_Name = "mdModernSubclassing"
'=========================================================================
'
' MST Project (c) 2019 by wqweto@gmail.com
'
' The Modern Subclassing Thunk (MST) for VB6
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
'Private Const MODULE_NAME As String = "mdModernSubclassing"

#Const ImplNoIdeProtection = (MST_NO_IDE_PROTECTION <> 0)
#Const ImplSelfContained = (MST_SELF_CONTAINED <> 0)

'=========================================================================
' API
'=========================================================================

Private Const SIGN_BIT                      As Long = &H80000000
Private Const PTR_SIZE                      As Long = 4
'--- for thunks
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcOrdinal As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#If Not ImplNoIdeProtection Then
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
#End If
#If ImplSelfContained Then
    Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
#End If

'=========================================================================
' Functions
'=========================================================================

Public Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As Object
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

Public Function InitSubclassingThunk(ByVal hWnd As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgepwEL4AV1aLdCQUg8YIgz4AdC+L+oHHFBK+AIvCBQgRvgCri8IFRBG+AKuLwgVUEb4Aq4vCBXwRvgCruQkAAADzpYHCFBK+AFJqGP9SEFqL+IvCq7gBAAAAqzPAq4tEJAyri3QkFKWlg+8YagBX/3IM/3cM/1IYi0QkGIk4Xl+4SBK+AC1wEL4AwhAAZpCLRCQIgzgAdSqDeAQAdSSBeAjAAAAAdRuBeAwAAABGdRKLVCQE/0IEi0QkDIkQM8DCDAC4AkAAgMIMAJCLVCQE/0IEi0IEwgQADx8Ai1QkBP9KBItCBHUYiwpS/3EM/3IM/1Eci1QkBIsKUv9RFDPAwgQAkFWL7ItVGP9CBItCEIXAdGiLCotBLIXAdDdS/9BaiUIIg/gBd1OFwHUJgX0MAwIAAHRGiwpS/1EwWoXAdTuLClJq8P9xJP9RKFqpAAAACHUoUjPAUFCNRCQEUI1EJARQ/3UU/3UQ/3UM/3UI/3IQ/1IUWVhahcl1E1KLCv91FP91EP91DP91CP9RIFpQUuhG////WF3CGACQ" ' 13.5.2020 17:56:48
    Const THUNK_SIZE    As Long = 472
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

Public Function TerminateSubclassingThunk(pSubclass As IUnknown, pObj As Object) As IUnknown
    If Not pSubclass Is Nothing Then
        Debug.Assert ThunkPrivateData(pSubclass, 2) = ObjPtr(pObj)
        ThunkPrivateData(pSubclass, 2) = 0
        Set pSubclass = Nothing
    End If
End Function

Public Function CallNextSubclassProc(pSubclass As IUnknown, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    #If pSubclass Then '--- touch args
    #End If
    CallNextSubclassProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
End Function

Public Function InitHookingThunk(ByVal idHook As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgepwEN4AV1aLdCQUg8YIgz4AdCqL+oHHTBLeAIvCBVQR3gCri8IFkBHeAKuLwgWgEd4AqzPAq7kJAAAA86WBwkwS3gBSahj/UhBai/iLwqu4AQAAAKszwKuri3QkFKWlg+8Yi0oM/0IMgWIM/wAAAI0Eyo0EyI1MiDTHAf80JLiJeQTHQQiJRCQEi8ItTBLeAAXEEd4AUMHgCAW4AAAAiUEMWMHoGAUA/+CQiUEQ/3QkEGoAUf90JBiLD/9RGIlHDItEJBiJOF5fuIAS3gAtcBDeAAUAFAAAwhAAi0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1FIsK/3IM/1Eci1QkBIsKUv9RFDPAwgQAkFWL7ItVCP9CBItCEIXAdFiLCotBLIXAdCpS/9BaiUIIg/gBd0OLClL/UTBahcB1OIsKUmrw/3Ek/1EoWqkAAAAIdSVSM8BQUI1EJARQjUQkBFD/dRT/dRD/dQz/chD/UhRZWFqFyXUTUosK/3UU/3UQ/3UM/3IM/1EgWlBS6Fr///9YXcIQAJA=" ' 13.5.2020 18:24:28
    Const THUNK_SIZE    As Long = 5648
    Static hThunk       As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitHookingThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        aParams(4) = GetProcAddress(GetModuleHandle("user32"), "SetWindowsHookExA")
        aParams(5) = GetProcAddress(GetModuleHandle("user32"), "UnhookWindowsHookEx")
        aParams(6) = GetProcAddress(GetModuleHandle("user32"), "CallNextHookEx")
        '--- for IDE protection
        Debug.Assert pvThunkIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitHookingThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, idHook, App.ThreadID, VarPtr(aParams(0)), VarPtr(InitHookingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Public Function TerminateHookingThunk(pHook As IUnknown, pObj As Object) As IUnknown
    If Not pHook Is Nothing Then
        Debug.Assert ThunkPrivateData(pHook, 2) = ObjPtr(pObj)
        ThunkPrivateData(pHook, 2) = 0
        Set pHook = Nothing
    End If
End Function

Public Function CallNextHookProc(pHook As IUnknown, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pHook)
    If lPtr <> 0 Then
        Call CopyMemory(lPtr, ByVal (lPtr Xor SIGN_BIT) + 12 Xor SIGN_BIT, PTR_SIZE)
    End If
    CallNextHookProc = CallNextHookEx(lPtr, nCode, wParam, lParam)
End Function

Public Function InitFireOnceTimerThunk(pObj As Object, ByVal pfnCallback As Long, Optional Delay As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgeogEQUAV1aLdCQUg8YIgz4AdCqL+oHHDBMFAIvCBSgSBQCri8IFZBIFAKuLwgV0EgUAqzPAq7kIAAAA86WBwgwTBQBSahj/UhBai/iLwqu4AQAAAKszwKuri3QkFKWlg+8Yi0IMSCX/AAAAUItKDDsMJHULWIsPV/9RFDP/62P/QgyBYgz/AAAAjQTKjQTIjUyIMIB5EwB101jHAf80JLiJeQTHQQiJRCQEi8ItDBMFAAWgEgUAUMHgCAW4AAAAiUEMWMHoGAUA/+CQiUEQiU8MUf90JBRqAGoAiw//URiJRwiLRCQYiTheX7g8EwUALSARBQAFABQAAMIQAGaQi0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1HYtCDMZAEwCLCv9yCGoA/1Eci1QkBIsKUv9RFDPAwgQAi1QkBItCEIXAdFuLCotBKIXAdCdS/9Bag/gBd0mLClL/USxahcB1PosKUmrw/3Eg/1EkWqkAAAAIdSuLClL/cghqAP9RHFr/QgQzwFBU/3IQ/1IUi1QkCMdCCAAAAABS6G////9YwhQADx8AjURAAQ==" ' 13.5.2020 18:59:12
    Const THUNK_SIZE    As Long = 5660
    Static hThunk       As Long
    Dim aParams(0 To 9) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitFireOnceTimerThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        aParams(4) = GetProcAddress(GetModuleHandle("user32"), "SetTimer")
        aParams(5) = GetProcAddress(GetModuleHandle("user32"), "KillTimer")
        '--- for IDE protection
        Debug.Assert pvThunkIdeOwner(aParams(6))
        If aParams(6) <> 0 Then
            aParams(7) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(8) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitFireOnceTimerThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, 0, Delay, VarPtr(aParams(0)), VarPtr(InitFireOnceTimerThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Public Function TerminateFireOnceTimerThunk(pTimer As IUnknown, pObj As Object) As IUnknown
    If Not pTimer Is Nothing Then
        Debug.Assert ThunkPrivateData(pTimer, 2) = ObjPtr(pObj)
        ThunkPrivateData(pTimer, 2) = 0
        Set pTimer = Nothing
    End If
End Function

Public Property Get ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long) As Long
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ThunkPrivateData, ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, PTR_SIZE)
    End If
End Property

Public Property Let ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long, ByVal lValue As Long)
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, lValue, PTR_SIZE)
    End If
End Property

Public Function InitCleanupThunk(ByVal hHandle As Long, sModuleName As String, sProcName As String) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgepQEDwBV1aLdCQUgz4AdCeL+oHHPBE8AYvCBcwQPAGri8IFCBE8AauLwgUYETwBq7kCAAAA86WBwjwRPAFSahD/Ugxai/iLwqu4AQAAAKuLRCQMq4tEJBCrg+8Qi0QkGIk4Xl+4UBE8AS1QEDwBwhAAkItEJAiDOAB1KoN4BAB1JIF4CMAAAAB1G4F4DAAAAEZ1EotUJAT/QgSLRCQMiRAzwMIMALgCQACAwgwAkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEdRL/cgj/UgyLVCQEiwpS/1EQM8DCBAAPHwA=" ' 25.3.2019 14:03:56
    Const THUNK_SIZE    As Long = 256
    Static hThunk       As Long
    Dim aParams(0 To 1) As Long
    Dim pfnCleanup      As Long
    Dim lSize           As Long
    
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitCleanupThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(0) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(1) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        #If ImplSelfContained Then
            pvThunkGlobalData("InitCleanupThunk") = hThunk
        #End If
    End If
    If Left$(sProcName, 1) = "#" Then
        pfnCleanup = GetProcByOrdinal(GetModuleHandle(sModuleName), Mid$(sProcName, 2))
    Else
        pfnCleanup = GetProcAddress(GetModuleHandle(sModuleName), sProcName)
    End If
    If pfnCleanup <> 0 Then
        lSize = CallWindowProc(hThunk, hHandle, pfnCleanup, VarPtr(aParams(0)), VarPtr(InitCleanupThunk))
        Debug.Assert lSize = THUNK_SIZE
    End If
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
