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

#Const ImplNoVBIDESupport = (MST_NO_VBIDE_SUPPORT <> 0)

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
#If Not ImplNoVBIDESupport Then
    Private Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadID As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
#End If

'=========================================================================
' Functions
'=========================================================================

Public Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As Object
    Const STR_THUNK     As String = "6AAAAABag+oFV4v6ge9QECkAgcekESkAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBApAIvCBTwRKQCri8IFUBEpAKuLwgVgESkAq4vCBYQRKQCri8IFjBEpAKuLwgWUESkAq4vCBZwRKQCri8IFpBEpALn5BwAAq4PABOL6i8dfgcJQECkAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 22.3.2019 10:17:01
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    Call CryptStringToBinary(STR_THUNK, Len(STR_THUNK), CRYPT_STRING_BASE64, hThunk, THUNK_SIZE)
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Public Function InitSubclassingThunk(ByVal hWnd As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Const STR_THUNK     As String = "6AAAAABag+oFgepwENQAV1aLdCQUg8YIgz4AdC+L+oHH4BHUAIvCBQQR1ACri8IFQBHUAKuLwgVQEdQAq4vCBXwR1ACruQkAAADzpYHC4BHUAFJqFP9SEFqL+IvCq7gBAAAAq4tEJAyri3QkFKWlg+8UagBX/3IM/3cI/1IYi0QkGIk4Xl+4FBLUAC1wENQAwhAAkItEJAiDOAB1KoN4BAB1JIF4CMAAAAB1G4F4DAAAAEZ1EotUJAT/QgSLRCQMiRAzwMIMALgCQACAwgwAkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfxiLClL/cQz/cgj/URyLVCQEiwpS/1EUM8DCBABmkFWL7ItVGIsKi0EshcB0J1L/0FqD+AF1N4sKUv9RMFqFwHUsiwpSavD/cST/UShaqQAAAAh1GTPAUFT/dRT/dRD/dQz/dQj/cgz/UhBY6xGLCv91FP91EP91DP91CP9RIF3CGAA=" ' 22.3.2019 10:17:35
    Const THUNK_SIZE    As Long = 420
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
            If InIde Then
                aParams(7) = hIdeOwner
                aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
                aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
                aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
            End If
        #End If
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
    Const STR_THUNK     As String = "6AAAAABag+oFgepwEIsAV1aLdCQUg8YIgz4AdCqL+oHHLBKLAIvCBVQRiwCri8IFkBGLAKuLwgWgEYsAqzPAq7kJAAAA86WBwiwSiwBSahT/UhBai/iLwqu4AQAAAKszwKuLdCQUpaWD7xSLSgz/QgyBYgz/AAAAjQTKjQTIjUyINMcB/zQkuIl5BMdBCIlEJASLwi0sEosABcgRiwBQweAIBbgAAACJQQxYwegYBQD/4JCJQRD/dCQQagBR/3QkGIsP/1EYiUcIi0QkGIk4Xl+4YBKLAC1wEIsABQAUAADCEACQi0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgSD+AB/FIsK/3II/1Eci1QkBIsKUv9RFDPAwgQAZpBVi+yLVQiLCotBLIXAdCdS/9Bag/gBdTSLClL/UTBahcB1KYsKUmrw/3Ek/1EoWqkAAAAIdRYzwFBU/3UU/3UQ/3UM/3IM/1IQWOsRiwr/dRT/dRD/dQz/cgj/USBdwhAADx8A" ' 22.3.2019 10:18:11
    Const THUNK_SIZE    As Long = 5616
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
        #If Not ImplNoVBIDESupport Then
            If InIde Then
                aParams(7) = hIdeOwner
                aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
                aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
                aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
            End If
        #End If
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
    Const STR_THUNK     As String = "6AAAAABag+oFgeogEYkAV1aLdCQUg8YIgz4AdCqL+oHH3BKJAIvCBQQSiQCri8IFQBKJAKuLwgVQEokAqzPAq7kIAAAA86WBwtwSiQBSahT/UhBai/iLwqu4AQAAAKszwKuLdCQUpaWD7xSLSgz/QgyBYgz/AAAAjQTKjQTIjUyIMMcB/zQkuIl5BMdBCIlEJASLwi3cEokABXgSiQBQweAIBbgAAACJQQxYwegYBQD/4JCJQRBR/3QkFGoAagCLD/9RGIlHCItEJBiJOF5fuAwTiQAtIBGJAAUAFAAAwhAADx8Ai0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgSD+AB/FosK/3IIagD/URyLVCQEiwpS/1EUM8DCBACLVCQEiwqLQSiFwHQzUv/QWoP4AXVJiwpS/1EsWoXAdT6LClJq8P9xIP9RJFqpAAAACHUriwpS/3IIagD/URxa/0IEM8BQVP9yDP9SEItUJAjHQggAAAAAUuh6////WMIUAGaQ" ' 22.3.2019 10:18:57
    Const THUNK_SIZE    As Long = 5612
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
        #If Not ImplNoVBIDESupport Then
            If InIde Then
                aParams(6) = hIdeOwner
                aParams(7) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
                aParams(8) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
                aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
            End If
        #End If
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
    Const STR_THUNK     As String = "6AAAAABag+oFgepQEPkAV1aLdCQUgz4AdCeL+oHHPBH5AIvCBcwQ+QCri8IFCBH5AKuLwgUYEfkAq7kCAAAA86WBwjwR+QBSahD/Ugxai/iLwqu4AQAAAKuLRCQMq4tEJBCrg+8Qi0QkGIk4Xl+4UBH5AC1QEPkAwhAAkItEJAiDOAB1KoN4BAB1JIF4CMAAAAB1G4F4DAAAAEZ1EotUJAT/QgSLRCQMiRAzwMIMALgCQACAwgwAkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfxL/cgj/UgyLVCQEiwpS/1EQM8DCBAA=" ' 22.3.2019 10:19:28
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

#If Not ImplNoVBIDESupport Then
Private Property Get InIde() As Boolean
    Debug.Assert pvSetTrue(InIde)
End Property

Private Function pvSetTrue(bValue As Boolean) As Boolean
    bValue = True
    pvSetTrue = True
End Function

Private Property Get hIdeOwner() As Long
    Call EnumThreadWindows(App.ThreadID, AddressOf pvEnumIdeOwner, VarPtr(hIdeOwner))
End Property

Private Function pvEnumIdeOwner(ByVal hWnd As Long, lResult As Long) As Long
    Dim sBuffer         As String
    
    sBuffer = String$(50, 0)
    Call GetClassName(hWnd, sBuffer, Len(sBuffer) - 1)
    If Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1) = "IDEOwner" Then
        lResult = hWnd
        Exit Function
    End If
    pvEnumIdeOwner = 1
End Function
#End If
