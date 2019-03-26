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

'=========================================================================
' API
'=========================================================================

Private Const SIGN_BIT                      As Long = &H80000000
'--- for thunks
Private Const MEM_COMMIT                    As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE        As Long = &H40
Private Const CRYPT_STRING_BASE64           As Long = 1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryA" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, pcbBinary As Long, Optional ByVal pdwSkip As Long, Optional ByVal pdwFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcAddressByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcOrdinal As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#If Not ImplNoIdeProtection Then
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

'=========================================================================
' Functions
'=========================================================================

Public Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As Object
    Const STR_THUNK     As String = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Public Function InitSubclassingThunk(ByVal hWnd As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
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
        '--- for IDE protection
        Debug.Assert pvGetIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
    End If
    lSize = CallWindowProc(hThunk, hWnd, 0, VarPtr(aParams(0)), VarPtr(InitSubclassingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Public Function CallNextSubclassProc(pSubclass As IUnknown, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    #If pSubclass Then '--- touch args
    #End If
    CallNextSubclassProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
End Function

Public Function InitHookingThunk(ByVal idHook As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Const STR_THUNK     As String = "6AAAAABag+oFgepwEMQAV1aLdCQUg8YIgz4AdCqL+oHHKBLEAIvCBVQRxACri8IFkBHEAKuLwgWgEcQAqzPAq7kJAAAA86WBwigSxABSahT/UhBai/iLwqu4AQAAAKszwKuLdCQUpaWD7xSLSgz/QgyBYgz/AAAAjQTKjQTIjUyINMcB/zQkuIl5BMdBCIlEJASLwi0oEsQABcQRxABQweAIBbgAAACJQQxYwegYBQD/4JCJQRD/dCQQagBR/3QkGIsP/1EYiUcIi0QkGIk4Xl+4XBLEAC1wEMQABQAUAADCEACQi0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1FIsK/3II/1Eci1QkBIsKUv9RFDPAwgQAkFWL7ItVCIsKi0EshcB0J1L/0FqD+AF1NIsKUv9RMFqFwHUpiwpSavD/cST/UShaqQAAAAh1FjPAUFT/dRT/dRD/dQz/cgz/UhBY6xGLCv91FP91EP91DP9yCP9RIF3CEAAPHwA=" ' 25.3.2019 14:02:50
    Const THUNK_SIZE    As Long = 5612
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
        aParams(4) = GetProcAddress(GetModuleHandle("user32"), "SetWindowsHookExA")
        aParams(5) = GetProcAddress(GetModuleHandle("user32"), "UnhookWindowsHookEx")
        aParams(6) = GetProcAddress(GetModuleHandle("user32"), "CallNextHookEx")
        '--- for IDE protection
        Debug.Assert pvGetIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
    End If
    lSize = CallWindowProc(hThunk, idHook, App.ThreadID, VarPtr(aParams(0)), VarPtr(InitHookingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Public Function CallNextHookProc(pHook As IUnknown, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pHook)
    If lPtr <> 0 Then
        Call CopyMemory(lPtr, ByVal (lPtr Xor SIGN_BIT) + 8 Xor SIGN_BIT, 4)
    End If
    CallNextHookProc = CallNextHookEx(lPtr, nCode, wParam, lParam)
End Function

Public Function InitFireOnceTimerThunk(pObj As Object, ByVal pfnCallback As Long, Optional Delay As Long) As IUnknown
    Const STR_THUNK     As String = "6AAAAABag+oFgeogEREBV1aLdCQUg8YIgz4AdCqL+oHHBBMRAYvCBSgSEQGri8IFZBIRAauLwgV0EhEBqzPAq7kIAAAA86WBwgQTEQFSahj/UhBai/iLwqu4AQAAAKszwKuri3QkFKWlg+8Yi0IMSCX/AAAAUItKDDsMJHULWIsPV/9RFDP/62P/QgyBYgz/AAAAjQTKjQTIjUyIMIB5EwB101jHAf80JLiJeQTHQQiJRCQEi8ItBBMRAQWgEhEBUMHgCAW4AAAAiUEMWMHoGAUA/+CQiUEQiU8MUf90JBRqAGoAiw//URiJRwiLRCQYiTheX7g0ExEBLSAREQEFABQAAMIQAGaQi0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1HYtCDMZAEwCLCv9yCGoA/1Eci1QkBIsKUv9RFDPAwgQAi1QkBIsKi0EohcB0J1L/0FqD+AF1SYsKUv9RLFqFwHU+iwpSavD/cSD/USRaqQAAAAh1K4sKUv9yCGoA/1EcWv9CBDPAUFT/chD/UhSLVCQIx0IIAAAAAFLodv///1jCFABmkA==" ' 25.3.2019 14:03:25
    Const THUNK_SIZE    As Long = 5652
    Static hThunk       As Long
    Dim aParams(0 To 9) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    If hThunk = 0 Then
        hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        aParams(4) = GetProcAddress(GetModuleHandle("user32"), "SetTimer")
        aParams(5) = GetProcAddress(GetModuleHandle("user32"), "KillTimer")
        '--- for IDE protection
        Debug.Assert pvGetIdeOwner(aParams(6))
        If aParams(6) <> 0 Then
            aParams(7) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(8) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
    End If
    lSize = CallWindowProc(hThunk, 0, Delay, VarPtr(aParams(0)), VarPtr(InitFireOnceTimerThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function

Property Get ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long) As Long
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ThunkPrivateData, ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, 4)
    End If
End Property

Property Let ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long, ByVal lValue As Long)
    Dim lPtr            As Long
    
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ByVal (lPtr Xor SIGN_BIT) + 8 + Index * 4 Xor SIGN_BIT, lValue, 4)
    End If
End Property

Public Function InitCleanupThunk(ByVal hHandle As Long, sModuleName As String, sProcName As String) As IUnknown
    Const STR_THUNK     As String = "6AAAAABag+oFgepQEDwBV1aLdCQUgz4AdCeL+oHHPBE8AYvCBcwQPAGri8IFCBE8AauLwgUYETwBq7kCAAAA86WBwjwRPAFSahD/Ugxai/iLwqu4AQAAAKuLRCQMq4tEJBCrg+8Qi0QkGIk4Xl+4UBE8AS1QEDwBwhAAkItEJAiDOAB1KoN4BAB1JIF4CMAAAAB1G4F4DAAAAEZ1EotUJAT/QgSLRCQMiRAzwMIMALgCQACAwgwAkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEdRL/cgj/UgyLVCQEiwpS/1EQM8DCBAAPHwA=" ' 25.3.2019 14:03:56
    Const THUNK_SIZE    As Long = 256
    Static hThunk       As Long
    Dim aParams(0 To 1) As Long
    Dim pfnCleanup      As Long
    Dim lSize           As Long
    
    If hThunk = 0 Then
        hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
        aParams(0) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(1) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
    End If
    pfnCleanup = GetProcAddress(GetModuleHandle(sModuleName), sProcName)
    If pfnCleanup <> 0 Then
        lSize = CallWindowProc(hThunk, hHandle, pfnCleanup, VarPtr(aParams(0)), VarPtr(InitCleanupThunk))
        Debug.Assert lSize = THUNK_SIZE
    End If
End Function

Private Function pvGetIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvGetIdeOwner = True
End Function
