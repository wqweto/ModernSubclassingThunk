//=========================================================================
//
// MST Project (c) 2019 by wqweto@gmail.com
//
// The Modern Subclassing Thunk (MST) for VB6
//
// This project is licensed under the terms of the MIT license
// See the LICENSE file in the project root for more information
//
//=========================================================================

//#define IMPL_ADDRESSOFMETHOD_THUNK
//#define IMPL_SUBCLASSING_THUNK
//#define IMPL_HOOKING_THUNK
//#define IMPL_FIREONCETIMER_THUNK
#define IMPL_CLEANUP_THUNK

#include <stdio.h>
#include <windows.h>
#include <commctrl.h>

#pragma comment(lib, "crypt32")
#pragma comment(lib, "comctl32")

LPWSTR __stdcall GetCurrentDateTime()
{
    static WCHAR szResult[50];
    SYSTEMTIME  st;
    DATE        dt;
    VARIANT     vdt = { VT_DATE, };
    VARIANT     vstr = { VT_EMPTY };

    GetLocalTime(&st);
    SystemTimeToVariantTime(&st, &dt);
    vdt.date = dt;
    VariantChangeType(&vstr, &vdt, 0, VT_BSTR);
    wcscpy_s(szResult, vstr.bstrVal);
    VariantClear(&vstr);
    return szResult;
}

int __stdcall EbMode()
{
    return 1;
}
int __stdcall EbIsResetting()
{
    return 0;
}

#ifdef IMPL_ADDRESSOFMETHOD_THUNK
enum InstanceData {
    m_pVtbl = 0,
    m_dwRefCnt = 4,
    m_pUnkVtbl = 8,
    m_pfnVirtualFree = 12,
    countof_DynamicMethods = 2048,
};

__declspec(naked, noinline)
void __stdcall AddressOfMethod(int hWnd, int uMsg, int wParam, int lParam)
{
    __asm {
__start:
        call    __next
__next:
        pop     edx
        sub     edx, 5
        push    edi
        // codegen dynamic methods
        mov     edi, edx
        sub     edi, offset __start
        add     edi, offset __DynamicMethods
        mov     eax, 0x042444FF
        mov     ecx, countof_DynamicMethods-7   // 7 methods from IDispatch
        rep     stosd
        mov     eax, edx
        shl     eax, 8
        add     eax, 0x000000b9
        stosd
        mov     eax, edx
        shr     eax, 24
        add     eax, 0x00818d00
        stosd
        mov     eax, 0x2b000000 + (countof_DynamicMethods >> 8)
        stosd
        mov     eax, 0x8b042444
        stosd
        mov     eax, 0x048b0849
        stosd
        mov     eax, 0x24548b81
        stosd
        mov     eax, dword ptr [esp+12]         // param uMsg
        shl     eax, 2
        add     eax, 0x33028908
        stosd
        mov     eax, dword ptr [esp+12]         // param uMsg
        shl     eax, 18
        add     eax, 0x0008c2c0
        stosd
        // init instance pVtbl, dwRefCnt and pfnVirtualFree members
        mov     dword ptr [edx+m_pVtbl], edi
        mov     dword ptr [edx+m_dwRefCnt], 1
        mov     eax, dword ptr [esp+8]          // param hWnd
        mov     eax, dword ptr [eax]            // pUnk->pVtbl
        mov     dword ptr [edx+m_pUnkVtbl], eax
        mov     eax, dword ptr [esp+16]         // param wParam
        mov     dword ptr [edx+m_pfnVirtualFree], eax
        // init vtable
        sub     edx, offset __start
        mov     eax, edx
        add     eax, offset __QI
        stosd
        mov     eax, edx
        add     eax, offset __AddRef
        stosd
        mov     eax, edx
        add     eax, offset __Release
        stosd
        mov     eax, edx
        add     eax, offset __GetTypeInfoCount
        stosd
        mov     eax, edx
        add     eax, offset __GetTypeInfo
        stosd
        mov     eax, edx
        add     eax, offset __GetIDsOfNames
        stosd
        mov     eax, edx
        add     eax, offset __Invoke
        stosd
        mov     eax, edx
        add     eax, offset __DynamicMethods
        mov     ecx, countof_DynamicMethods-7
__loop1:
        stosd
        add     eax, 4
        loop    __loop1
        mov     eax, edi
        pop     edi
        add     edx, offset __start
        mov     ecx, dword ptr [esp+16]         // param lParam
        mov     dword ptr [ecx], edx            // *lParam = this ptr
        sub     eax, edx
        ret     16

        align   4
__QI:
        mov     edx, dword ptr [esp+4]          // param this
        inc     dword ptr [edx+m_dwRefCnt]
        mov     eax, dword ptr [esp+12]         // param ppvObject
        mov     dword ptr [eax], edx            // *ppvObject = this ptr
        xor     eax, eax
        ret     12

        align   4
__AddRef:
        mov     edx, dword ptr [esp+4]          // param this
        inc     dword ptr [edx+m_dwRefCnt]
        mov     eax, dword ptr [edx+m_dwRefCnt]
        ret     4

        align   4
__Release:
        mov     edx, dword ptr [esp+4]          // param this
        dec     dword ptr [edx+m_dwRefCnt]
        mov     eax, dword ptr [edx+m_dwRefCnt]
        cmp     eax, 0
        jle     __free
        ret     4
__free:
        pop     ecx                             // pop ret addr
        pop     edx                             // this param
        mov     eax, dword ptr [edx+m_pfnVirtualFree]
        push    0x8000                          // MEM_RELEASE
        push    0
        push    edx                             // this
        push    ecx                             // push ret addr
        jmp     eax

        align   4
__GetTypeInfoCount:
        mov     eax, 0x80004001                 // E_NOTIMPL
        ret     8

        align   4
__GetTypeInfo:
        mov     eax, 0x80004001                 // E_NOTIMPL
        ret     16

        align   4
__GetIDsOfNames:
        mov     eax, 0x80004001                 // E_NOTIMPL
        ret     24

        align   4
__Invoke:
        mov     eax, 0x80004001                 // E_NOTIMPL
        ret     36

        align   4
__DynamicMethods:
        //inc     dword ptr [esp+8]
        //inc     dword ptr [esp+8]
        //...
        mov     ecx, 0xDEADBEEF                 // this ptr
        lea     eax, [ecx+countof_DynamicMethods]
        sub     eax, dword ptr [esp+4]          // eax = MethodIndex
        mov     ecx, dword ptr [ecx+m_pUnkVtbl] // this->pUnkVtbl
        mov     eax, dword ptr [ecx+4*eax]      // eax = this->pUnkVtbl[MethodIndex]
        mov     edx, dword ptr [esp+24]
        mov     dword ptr [edx], eax            // *pResult = eax
        xor     eax, eax
        ret     24                              // 4*uMsg + 8 (8 = this+pResult)
    }
}

#define THUNK_SIZE 340
// ((char *)main - (char *)AddressOfMethod)

class ICustom : public IDispatch
{
public:
    virtual HRESULT __stdcall FormSubclassProc(HWND hWnd, UINT uMsg, LPARAM lParam, WPARAM wParam,
                                                       LRESULT *pRetVal) = 0;
};

HRESULT __stdcall SubclassProc(void *self, HWND hWnd, UINT uMsg, LPARAM lParam, WPARAM wParam, LRESULT *pRetVal);

void main()
{
    CoInitialize(0);
    //AddressOfMethod(0,0,0,0);
    void *vtbl[] = { 0, 0, 0, 0, 0, 0, 0, SubclassProc };
    void *pObj[] = { vtbl };
    void *hThunk = VirtualAlloc(0, 0x10000, MEM_COMMIT, PAGE_EXECUTE_READWRITE);
    printf("hThunk=%p\nTHUNK_SIZE=%d\n", hThunk, THUNK_SIZE);
    memcpy(hThunk, AddressOfMethod, THUNK_SIZE);
    ICustom *pUnk = 0;
    DWORD dwSize = CallWindowProc((WNDPROC)hThunk, (HWND)pObj, 4, (WPARAM)VirtualFree, (LPARAM)&pUnk);
    printf("dwSize=%d\n", dwSize);
    printf("SubclassProc=%p\n", SubclassProc);
    LRESULT lResult;
    pUnk->FormSubclassProc(0, 0, 0, 0, &lResult);
    printf("pCustom->FormSubclassProc=%p\n", lResult);
    printf("pUnk->AddRef=%d\n", pUnk->AddRef());
    IID iid;
    printf("pDisp->GetTypeInfoCount=%08X\n", pUnk->GetTypeInfoCount(0));
    printf("pDisp->GetTypeInfo=%08X\n", pUnk->GetTypeInfo(0, 0, 0));
    printf("pDisp->GetIDsOfNames=%08X\n", pUnk->GetIDsOfNames(iid, 0, 0, 0, 0));
    printf("pDisp->Invoke=%08X\n", pUnk->Invoke(0, iid, 0, 0, 0, 0, 0, 0));
    printf("pUnk->Release=%d\n", pUnk->Release());
    printf("pUnk->Release=%d\n", pUnk->Release());
    WCHAR szBuffer[1000];
    DWORD dwBufSize = _countof(szBuffer);
    CryptBinaryToString((BYTE *)AddressOfMethod, THUNK_SIZE, CRYPT_STRING_BASE64, szBuffer, &dwBufSize);
    for(int i = 0, j = 0; (szBuffer[j] = szBuffer[i]) != 0; )
        ++i, j += (szBuffer[j] != '\r' && szBuffer[j] != '\n');
    printf("Const STR_THUNK     As String = \"%S\" ' %S\n", szBuffer, GetCurrentDateTime());
}

HRESULT __stdcall SubclassProc(void *self, HWND hWnd, UINT uMsg, LPARAM lParam, WPARAM wParam, LRESULT *pRetVal)
{
    return S_OK;
}
#endif

#ifdef IMPL_SUBCLASSING_THUNK
struct InParams {
    void *pCallbackThis;
    void *pfnCallback;
    void *pfnCoTaskMemAlloc;
    void *pfnCoTaskMemFree;
    void *pfnSetWindowSubclass;
    void *pfnRemoveWindowSubclass;
    void *pfnDefSubclassProc;
    int hIdeOwner;
    void *pfnGetWindowLong;
    void *pfnEbMode;
    void *pfnEbIsResetting;
};

enum InstanceData {
    m_pVtbl = 0,
    m_dwRefCnt = 4,
    m_hWnd = 8,
    m_pCallbackThis = 12,
    m_pfnCallback = 16,
    sizeof_InstanceData = 20,
};
enum ThunkData {
    t_pfnQI = 0,
    t_pfnAddRef = 4,
    t_pfnRelease = 8,
    t_pfnSubclassProc = 12,
    t_pfnCoTaskMemAlloc = 16,
    t_pfnCoTaskMemFree = 20,
    t_pfnSetWindowSubclass = 24,
    t_pfnRemoveWindowSubclass = 28,
    t_pfnDefSubclassProc = 32,
    t_hIdeOwner = 36,
    t_pfnGetWindowLong = 40,
    t_pfnEbMode = 44,
    t_pfnEbIsResetting = 48,
    sizeof_ThunkData = 52,
    countof_pfn = 9
};

__declspec(naked, noinline)
void __stdcall SubclassingThunk(int hWnd, int uMsg, int wParam, int lParam)
{
    __asm {
__start:
        call    __next
__next:
        pop     edx
        sub     edx, 5
        sub     edx, offset __start
        push    edi
        push    esi

        // init thunk data
        mov     esi, dword ptr [esp+20]         // param wParam
        add     esi, 8
        cmp     dword ptr [esi], 0              // wParam->pfnCoTaskMemAlloc
        jz      __skip_init_thunk_data
        mov     edi, edx
        add     edi, offset __vtable
        mov     eax, edx
        add     eax, offset __QI
        stosd                                   // vtbl->pfnQI
        mov     eax, edx
        add     eax, offset __AddRef
        stosd                                   // vtbl->pfnAddRef
        mov     eax, edx
        add     eax, offset __Release
        stosd                                   // vtbl->pfnRelease
        mov     eax, edx
        add     eax, offset __SubclassProc
        stosd                                   // vtbl->pfnSubclassProc
        mov     ecx, countof_pfn
        rep     movsd                           // vtbl->pfn*
__skip_init_thunk_data:

        // this = CoTaskMemAlloc(sizeof_InstanceData)
        add     edx, offset __vtable
        push    edx
        push    sizeof_InstanceData
        call    dword ptr [edx+t_pfnCoTaskMemAlloc]
        pop     edx
        mov     edi, eax                        // edi = this ptr

        // init instance members
        mov     eax, edx
        stosd                                   // this->pVtbl
        mov     eax, 1
        stosd                                   // this->dwRefCnt
        mov     eax, dword ptr [esp+12]         // param hWnd
        stosd                                   // this->hWnd
        mov     esi, dword ptr [esp+20]         // param wParam
        movsd                                   // wParam->pCallbackThis
        movsd                                   // wParam->pfnCallback
        sub     edi, sizeof_InstanceData        // rewind this ptr

        // invoke SetWindowSubclass(hWnd, __SubclassProc, this, 0)
        push    0                               // 0 (for dwRefData)
        push    edi                             // this ptr (for uIdSubclass)
        push    dword ptr [edx+t_pfnSubclassProc] // vtbl->pfnSubclassProc (for pfnSubclass)
        push    dword ptr [edi+m_hWnd]          // this->hWnd (for hWnd)
        call    dword ptr [edx+t_pfnSetWindowSubclass]

        // retval & epilog
        mov     eax, dword ptr [esp+24]         // param lParam
        mov     dword ptr [eax], edi            // *lParam = this
        pop     esi
        pop     edi
        mov     eax, offset __end
        sub     eax, offset __start
        ret     16

        align   4
__QI:
        mov     eax, dword ptr [esp+8]          // eax = param riid
        cmp     dword ptr [eax+0], 0
        jne     __nointerface
        cmp     dword ptr [eax+4], 0
        jne     __nointerface
        cmp     dword ptr [eax+8], 0xC0
        jne     __nointerface
        cmp     dword ptr [eax+12], 0x46000000
        jne     __nointerface
        mov     edx, dword ptr [esp+4]          // edx = param this
        inc     dword ptr [edx+m_dwRefCnt]
        mov     eax, dword ptr [esp+12]         // eax = param ppvObject
        mov     dword ptr [eax], edx            // *ppvObject = this
        xor     eax, eax
        ret     12
__nointerface:
        mov     eax, 0x80004002                 // E_NOINTERFACE
        ret     12

        align   4
__AddRef:
        mov     edx, dword ptr [esp+4]          // edx = param this
        inc     dword ptr [edx+m_dwRefCnt]
        mov     eax, dword ptr [edx+m_dwRefCnt]
        ret     4

        align   4
__Release:
        mov     edx, dword ptr [esp+4]          // edx = param this
        dec     dword ptr [edx+m_dwRefCnt]
        mov     eax, dword ptr [edx+m_dwRefCnt]
        jnz     __exit_Release
        
        // invoke RemoveWindowSubclass(this->hWnd, __SubclassProc, this)
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    edx                             // this ptr (for uIdSubclass)
        push    dword ptr [ecx+t_pfnSubclassProc] // __SubclassProc (for pfnSubclass)
        push    dword ptr [edx+m_hWnd]          // this->hWnd (for hWnd)
        call    dword ptr [ecx+t_pfnRemoveWindowSubclass] // vtbl->pfnRemoveWindowSubclass
        
        // invoke CoTaskMemFree(this)
        mov     edx, dword ptr [esp+4]          // edx = param this
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    edx
        call    dword ptr [ecx+t_pfnCoTaskMemFree] // vtbl->pfnCoTaskMemFree
        xor     eax, eax
__exit_Release:
        ret     4

        align   4
__SubclassProc:
        push    ebp                             
        mov     ebp, esp
        mov     edx, dword ptr [ebp+24]         // this ptr = param uIdSubclass
        // if EbMode() > 1 goto __skip_callback
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        mov     eax, dword ptr [ecx+t_pfnEbMode]
        test    eax, eax
        jz      __call_callback
        push    edx
        call    eax
        pop     edx                             // restore this ptr
        cmp     eax, 1
        ja      __call_def_subclass
        // if EbIsResetting() goto __skip_callback
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    edx
        call    dword ptr [ecx+t_pfnEbIsResetting]
        pop     edx                             // restore this ptr
        test    eax, eax
        jnz     __call_def_subclass
        // invoke GetWindowLong(this->hIdeOwner, GWL_STYLE)
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    edx
        push    GWL_STYLE
        push    dword ptr [ecx+t_hIdeOwner]
        call    dword ptr [ecx+t_pfnGetWindowLong]
        pop     edx
        test    eax, WS_DISABLED
        jnz     __call_def_subclass
__call_callback:
        xor     eax, eax
        push    eax
        push    esp                             // pRetVal
        push    dword ptr [ebp+20]              // param lParam
        push    dword ptr [ebp+16]              // param wParam
        push    dword ptr [ebp+12]              // param uMsg
        push    dword ptr [ebp+8]               // param hWnd
        push    dword ptr [edx+m_pCallbackThis]
        call    dword ptr [edx+m_pfnCallback]
        pop     eax
        jmp     __exit_SubclassProc
__call_def_subclass:
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    dword ptr [ebp+20]              // param lParam
        push    dword ptr [ebp+16]              // param wParam
        push    dword ptr [ebp+12]              // param uMsg
        push    dword ptr [ebp+8]               // param hWnd
        call    dword ptr [ecx+t_pfnDefSubclassProc]
__exit_SubclassProc:
        pop     ebp
        ret     24
        align   4
__vtable:
        lea     eax,[eax+eax*2+1]               // t_pfnQI
        lea     eax,[eax+eax*2+1]               // t_pfnAddRef
        lea     eax,[eax+eax*2+1]               // t_pfnRelease
        lea     eax,[eax+eax*2+1]               // t_pfnSubclassProc
        lea     eax,[eax+eax*2+1]               // t_pfnCoTaskMemAlloc
        lea     eax,[eax+eax*2+1]               // t_pfnCoTaskMemFree
        lea     eax,[eax+eax*2+1]               // t_pfnSetWindowSubclass
        lea     eax,[eax+eax*2+1]               // t_pfnRemoveWindowSubclass
        lea     eax,[eax+eax*2+1]               // t_pfnDefSubclassProc
        lea     eax,[eax+eax*2+1]               // t_hIdeOwner
        lea     eax,[eax+eax*2+1]               // t_pfnGetWindowLong
        lea     eax,[eax+eax*2+1]               // t_pfnEbMode
        lea     eax,[eax+eax*2+1]               // t_pfnEbIsResetting
__end:
    }
}

#define THUNK_SIZE ((char *)main - (char *)SubclassingThunk)

HRESULT __stdcall SubclassProc(void *self, HWND hWnd, UINT uMsg, LPARAM lParam, WPARAM wParam, LRESULT *pRetVal);

void main()
{
    CoInitialize(0);
    LoadLibrary(L"comctl32");
    void *vtbl[] = { SubclassProc };
    void *pObj[] = { vtbl };
    InParams p = { 0 };
    p.pCallbackThis = pObj;
    p.pfnCallback = SubclassProc;
    p.pfnCoTaskMemAlloc = GetProcAddress(GetModuleHandle(L"ole32"), "CoTaskMemAlloc");
    p.pfnCoTaskMemFree = GetProcAddress(GetModuleHandle(L"ole32"), "CoTaskMemFree");
    p.pfnSetWindowSubclass = GetProcAddress(GetModuleHandle(L"comctl32"), "SetWindowSubclass");
    p.pfnRemoveWindowSubclass = GetProcAddress(GetModuleHandle(L"comctl32"), "RemoveWindowSubclass");
    p.pfnDefSubclassProc = GetProcAddress(GetModuleHandle(L"comctl32"), "DefSubclassProc");
    p.hIdeOwner = 0x1C520D8C;
    p.pfnGetWindowLong = GetProcAddress(GetModuleHandle(L"user32"), "GetWindowLongA");
    p.pfnEbMode = EbMode;
    p.pfnEbIsResetting = EbIsResetting;
    HWND hWnd = CreateWindowEx(0, L"STATIC", L"Test", 0, 0, 0, 0, 0, 0, 0, 0, 0);
    printf("hWnd=%p\n", hWnd);
    void *hThunk = VirtualAlloc(0, 0x10000, MEM_COMMIT, PAGE_EXECUTE_READWRITE);
    printf("hThunk=%p\nTHUNK_SIZE=%d\n", hThunk, THUNK_SIZE);
    memcpy(hThunk, SubclassingThunk, THUNK_SIZE);
    IUnknown *pUnk = 0;
    DWORD dwSize = CallWindowProc((WNDPROC)hThunk, hWnd, 0, (WPARAM)&p, (LPARAM)&pUnk);
    printf("dwSize=%d\nsizeof_InstanceData=%d\n", dwSize, sizeof_InstanceData);
    SendMessage(hWnd, WM_USER, 0, 0);
    printf("pUnk->AddRef=%d\n", pUnk->AddRef());
    printf("pUnk->Release=%d\n", pUnk->Release());
    printf("pUnk->Release=%d\n", pUnk->Release());
    WCHAR szBuffer[1000];
    DWORD dwBufSize = _countof(szBuffer);
    CryptBinaryToString((BYTE *)SubclassingThunk, dwSize - sizeof_ThunkData, CRYPT_STRING_BASE64, szBuffer, &dwBufSize);
    for(int i = 0, j = 0; (szBuffer[j] = szBuffer[i]) != 0; )
        ++i, j += (szBuffer[j] != '\r' && szBuffer[j] != '\n');
    printf("Const STR_THUNK     As String = \"%S\" ' %S\n", szBuffer, GetCurrentDateTime());
}

HRESULT __stdcall SubclassProc(void *self, HWND hWnd, UINT uMsg, LPARAM lParam, WPARAM wParam, LRESULT *pRetVal)
{
    printf("uMsg=%04X\n", uMsg);
    *pRetVal = DefSubclassProc(hWnd, uMsg, lParam, wParam);
    *pRetVal = 0x123;
    return S_OK;
}
#endif

#ifdef IMPL_HOOKING_THUNK
struct InParams {
    void *pCallbackThis;
    void *pfnCallback;
    void *pfnCoTaskMemAlloc;
    void *pfnCoTaskMemFree;
    void *pfnSetWindowsHookEx;
    void *pfnUnhookWindowsHookEx;
    void *pfnCallNextHookEx;
    int hIdeOwner;
    void *pfnGetWindowLong;
    void *pfnEbMode;
    void *pfnEbIsResetting;
};

enum InstanceData {
    m_pVtbl = 0,
    m_dwRefCnt = 4,
    m_hHook = 8,
    m_pCallbackThis = 12,
    m_pfnCallback = 16,
    sizeof_InstanceData = 20,
};
enum ThunkData {
    t_pfnQI = 0,
    t_pfnAddRef = 4,
    t_pfnRelease = 8,
    t_PushParamIdx = 12,
    t_pfnCoTaskMemAlloc = 16,
    t_pfnCoTaskMemFree = 20,
    t_pfnSetWindowsHookEx = 24,
    t_pfnUnhookWindowsHookEx = 28,
    t_pfnCallNextHookEx = 32,
    t_hIdeOwner = 36,
    t_pfnGetWindowLong = 40,
    t_pfnEbMode = 44,
    t_pfnEbIsResetting = 48,
    sizeof_ThunkData = 52,
    countof_pfn = 9
};
enum PushParamThunkData {
    sizeof_PushParamThunk = 20,
    countof_PushParamThunk = 256 // power of 2
};

__declspec(naked, noinline)
void __stdcall HookingThunk(int hWnd, int uMsg, int wParam, int lParam)
{
    __asm {
__start:
        call    __next
__next:
        pop     edx
        sub     edx, 5
        sub     edx, offset __start
        push    edi
        push    esi

        // init thunk data
        mov     esi, dword ptr [esp+20]         // param wParam
        add     esi, 8
        cmp     dword ptr [esi], 0              // wParam->pfnCoTaskMemAlloc
        jz      __skip_init_thunk_data
        mov     edi, edx
        add     edi, offset __vtable
        mov     eax, edx
        add     eax, offset __QI
        stosd                                   // vtbl->pfnQI
        mov     eax, edx
        add     eax, offset __AddRef
        stosd                                   // vtbl->pfnAddRef
        mov     eax, edx
        add     eax, offset __Release
        stosd                                   // vtbl->pfnRelease
        xor     eax, eax
        stosd                                   // vtbl->PushParamIdx
        mov     ecx, countof_pfn
        rep     movsd                           // vtbl->pfn*
__skip_init_thunk_data:

        // this = CoTaskMemAlloc(sizeof_InstanceData)
        add     edx, offset __vtable
        push    edx
        push    sizeof_InstanceData
        call    dword ptr [edx+t_pfnCoTaskMemAlloc]
        pop     edx
        mov     edi, eax                        // edi = this ptr

        // init instance members
        mov     eax, edx
        stosd                                   // this->pVtbl
        mov     eax, 1
        stosd                                   // this->dwRefCnt
        xor     eax, eax
        stosd                                   // this->hHook
        mov     esi, dword ptr [esp+20]         // param wParam
        movsd                                   // wParam->pCallbackThis
        movsd                                   // wParam->pfnCallback
        sub     edi, sizeof_InstanceData        // rewind this ptr

        // init pfnPushParamThunk
        mov     ecx, dword ptr [edx+t_PushParamIdx]
        inc     dword ptr [edx+t_PushParamIdx]
        and     dword ptr [edx+t_PushParamIdx], countof_PushParamThunk - 1
        lea     eax, [edx + 8*ecx]
        lea     eax, [eax + 8*ecx]
        lea     ecx, [eax + 4*ecx + sizeof_ThunkData] // ecx = pfnPushParamThunk
        mov     dword ptr [ecx+0], 0xB82434FF
        mov     dword ptr [ecx+4], edi
        mov     dword ptr [ecx+8], 0x04244489
        mov     eax, edx
        sub     eax, offset __vtable
        add     eax, offset __HookProc
        push    eax
        shl     eax, 8
        add     eax, 0x000000B8
        mov     dword ptr [ecx+12], eax
        pop     eax
        shr     eax, 24
        add     eax, 0x90E0FF00
        mov     dword ptr [ecx+16], eax

        // this->hHook = SetWindowsHookEx(idHook, pfnPushParamThunk, 0, dwThreadId)
        push    dword ptr [esp+16]              // param wMsg (for dwThreadId)
        push    0                               // for hMod
        push    ecx                             // pfnPushParamThunk (for lpfn)
        push    dword ptr [esp+24]              // param hWnd (for idHook)
        mov     ecx, dword ptr [edi]            // ecx = this->pVtbl
        call    dword ptr [ecx+t_pfnSetWindowsHookEx]
        mov     dword ptr [edi+m_hHook], eax

        // retval & epilog
        mov     eax, dword ptr [esp+24]         // param lParam
        mov     dword ptr [eax], edi            // *lParam = this
        pop     esi
        pop     edi
        mov     eax, offset __PushParamThunk
        sub     eax, offset __start
        add     eax, countof_PushParamThunk * sizeof_PushParamThunk
        ret     16

        align   4
__QI:
        mov     eax, dword ptr [esp+8]          // param riid
        cmp     dword ptr [eax+0], 0
        jne     __nointerface
        cmp     dword ptr [eax+4], 0
        jne     __nointerface
        cmp     dword ptr [eax+8], 0xC0
        jne     __nointerface
        cmp     dword ptr [eax+12], 0x46000000
        jne     __nointerface
        mov     edx, dword ptr [esp+4]          // param this
        inc     dword ptr [edx+m_dwRefCnt]
        mov     eax, dword ptr [esp+12]         // param ppvObject
        mov     dword ptr [eax], edx            // *ppvObject = this
        xor     eax, eax
        ret     12
__nointerface:
        mov     eax, 0x80004002                 // E_NOINTERFACE
        ret     12

        align   4
__AddRef:
        mov     edx, dword ptr [esp+4]          // param this
        inc     dword ptr [edx+m_dwRefCnt]
        mov     eax, dword ptr [edx+m_dwRefCnt]
        ret     4

        align   4
__Release:
        mov     edx, dword ptr [esp+4]          // param this
        dec     dword ptr [edx+m_dwRefCnt]
        mov     eax, dword ptr [edx+m_dwRefCnt]
        jnz     __exit_Release
        
        // invoke UnhookWindowsHookEx(this->hHook)
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    dword ptr [edx+m_hHook]         // this->hHook (for hhk)
        call    dword ptr [ecx+t_pfnUnhookWindowsHookEx]
        
        // invoke CoTaskMemFree(this)
        mov     edx, dword ptr [esp+4]          // param this
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    edx
        call    dword ptr [ecx+t_pfnCoTaskMemFree]
        xor     eax, eax
__exit_Release:
        ret     4

        align   4
__HookProc:
        push    ebp
        mov     ebp, esp
        mov     edx, dword ptr [ebp+8]          // this ptr
        // if EbMode() > 1 goto __skip_callback
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        mov     eax, dword ptr [ecx+t_pfnEbMode]
        test    eax, eax
        jz      __call_callback
        push    edx
        call    eax                             // this->pfnEbMode
        pop     edx
        cmp     eax, 1                          // 1 = running
        ja      __skip_callback
        // if EbIsResetting() goto __skip_callback
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    edx
        call    dword ptr [ecx+t_pfnEbIsResetting]
        pop     edx                             // restore this ptr
        test    eax, eax
        jnz      __skip_callback
        // invoke GetWindowLong(this->hIdeOwner, GWL_STYLE)
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    edx
        push    GWL_STYLE
        push    dword ptr [ecx+t_hIdeOwner]
        call    dword ptr [ecx+t_pfnGetWindowLong]
        pop     edx
        test    eax, WS_DISABLED
        jnz     __skip_callback
__call_callback:
        xor     eax, eax
        push    eax
        push    esp                             // pRetVal
        push    dword ptr [ebp+20]              // param lParam
        push    dword ptr [ebp+16]              // param wParam
        push    dword ptr [ebp+12]              // param nCode
        push    dword ptr [edx+m_pCallbackThis]
        call    dword ptr [edx+m_pfnCallback]
        pop     eax                             // RetVal
        jmp     __exit_HookProc
__skip_callback:
        // invoke CallNextHookEx(hHook, nCode, wParam, lParam)
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    dword ptr [ebp+20]              // param lParam
        push    dword ptr [ebp+16]              // param wParam
        push    dword ptr [ebp+12]              // param nCode
        push    dword ptr [edx+m_hHook]
        call    dword ptr [ecx+t_pfnCallNextHookEx]
__exit_HookProc:
        pop     ebp
        ret     16
        align   4
__vtable:
        lea     eax,[eax+eax*2+1]               // t_pfnQI
        lea     eax,[eax+eax*2+1]               // t_pfnAddRef
        lea     eax,[eax+eax*2+1]               // t_pfnRelease
        lea     eax,[eax+eax*2+1]               // t_PushParamIdx
        lea     eax,[eax+eax*2+1]               // t_pfnCoTaskMemAlloc
        lea     eax,[eax+eax*2+1]               // t_pfnCoTaskMemFree
        lea     eax,[eax+eax*2+1]               // t_pfnSetWindowsHookEx
        lea     eax,[eax+eax*2+1]               // t_pfnUnhookWindowsHookEx
        lea     eax,[eax+eax*2+1]               // t_pfnCallNextHookEx
        lea     eax,[eax+eax*2+1]               // t_hIdeOwner
        lea     eax,[eax+eax*2+1]               // t_pfnGetWindowLong
        lea     eax,[eax+eax*2+1]               // t_pfnEbMode
        lea     eax,[eax+eax*2+1]               // t_pfnEbIsResetting
__PushParamThunk:
        //push    dword ptr [esp]
        //mov     eax, 0xDEADBEEF
        //mov     dword ptr [esp+4], eax
        //mov     eax, 0xDEADBEEF
        //jmp     eax
    }
}

#define THUNK_SIZE ((char *)main - (char *)HookingThunk)

HRESULT __stdcall HookProc(void *self, int nCode, WPARAM wParam, LPARAM lParam, LRESULT *pRetVal);

void main()
{
    CoInitialize(0);
    void *vtbl[] = { HookProc };
    void *pObj[] = { vtbl };
    InParams p = { 0 };
    p.pCallbackThis = pObj;
    p.pfnCallback = HookProc;
    p.pfnCoTaskMemAlloc = GetProcAddress(GetModuleHandle(L"ole32"), "CoTaskMemAlloc");
    p.pfnCoTaskMemFree = GetProcAddress(GetModuleHandle(L"ole32"), "CoTaskMemFree");
    p.pfnSetWindowsHookEx = GetProcAddress(GetModuleHandle(L"user32"), "SetWindowsHookExA");
    p.pfnUnhookWindowsHookEx = GetProcAddress(GetModuleHandle(L"user32"), "UnhookWindowsHookEx");
    p.pfnCallNextHookEx = GetProcAddress(GetModuleHandle(L"user32"), "CallNextHookEx");
    p.hIdeOwner = 0x1C520D8C;
    p.pfnGetWindowLong = GetProcAddress(GetModuleHandle(L"user32"), "GetWindowLongA");
    p.pfnEbMode = EbMode;
    p.pfnEbIsResetting = EbIsResetting;
    void *hThunk = VirtualAlloc(0, 0x10000, MEM_COMMIT, PAGE_EXECUTE_READWRITE);
    printf("hThunk=%p\nTHUNK_SIZE=%d\n", hThunk, THUNK_SIZE);
    memcpy(hThunk, HookingThunk, THUNK_SIZE);
    IUnknown *pUnk = 0;
    printf("GetCurrentThreadId()=%04X\n", GetCurrentThreadId());
    DWORD dwSize = CallWindowProc((WNDPROC)hThunk, (HWND)WH_CALLWNDPROC, GetCurrentThreadId(), (WPARAM)&p, (LPARAM)&pUnk);
    printf("dwSize=%d\nsizeof_InstanceData=%d\n", dwSize, sizeof_InstanceData);
    dwSize -= countof_PushParamThunk * sizeof_PushParamThunk;
    printf("offset __PushParamThunk=%d\n", dwSize);
    dwSize -= sizeof_ThunkData;
    printf("offset __vtable=%d\n", dwSize);
    HWND hWnd = CreateWindowEx(0, L"STATIC", L"Test", 0, 0, 0, 0, 0, 0, 0, 0, 0);
    ShowWindow(hWnd, SW_SHOWMAXIMIZED);
    ShowWindow(hWnd, SW_HIDE);
    printf("pUnk->AddRef=%d\n", pUnk->AddRef());
    printf("pUnk->Release=%d\n", pUnk->Release());
    printf("pUnk->Release=%d\n", pUnk->Release());
    WCHAR szBuffer[1000];
    DWORD dwBufSize = _countof(szBuffer);
    CryptBinaryToString((BYTE *)HookingThunk, dwSize, CRYPT_STRING_BASE64, szBuffer, &dwBufSize);
    for(int i = 0, j = 0; (szBuffer[j] = szBuffer[i]) != 0; )
        ++i, j += (szBuffer[j] != '\r' && szBuffer[j] != '\n');
    printf("Const STR_THUNK     As String = \"%S\" ' %S\n", szBuffer, GetCurrentDateTime());
}

HRESULT __stdcall HookProc(void *self, int nCode, WPARAM wParam, LPARAM lParam, LRESULT *pRetVal)
{
    if (nCode == HC_ACTION) {
        CWPSTRUCT *cwp = (CWPSTRUCT *)lParam;
        printf("hwnd=%d, message=%d, wParam=%08X, lParam=%08X\n", cwp->hwnd, cwp->message, cwp->wParam, cwp->lParam);
    }
    *pRetVal = CallNextHookEx((HHOOK)((int *)self)[2], nCode, wParam, lParam);
    return S_OK;
}
#endif

#ifdef IMPL_FIREONCETIMER_THUNK
struct InParams {
    void *pCallbackThis;
    void *pfnCallback;
    void *pfnCoTaskMemAlloc;
    void *pfnCoTaskMemFree;
    void *pfnSetTimer;
    void *pfnKillTimer;
    int hIdeOwner;
    void *pfnGetWindowLong;
    void *pfnEbMode;
    void *pfnEbIsResetting;
};

enum InstanceData {
    m_pVtbl = 0,
    m_dwRefCnt = 4,
    m_dwTimerID = 8,
    m_pfnPushParamThunk = 12,
    m_pCallbackThis = 16,
    m_pfnCallback = 20,
    sizeof_InstanceData = 24,
};
enum ThunkData {
    t_pfnQI = 0,
    t_pfnAddRef = 4,
    t_pfnRelease = 8,
    t_PushParamIdx = 12,
    t_pfnCoTaskMemAlloc = 16,
    t_pfnCoTaskMemFree = 20,
    t_pfnSetTimer = 24,
    t_pfnKillTimer = 28,
    t_hIdeOwner = 32,
    t_pfnGetWindowLong = 36,
    t_pfnEbMode = 40,
    t_pfnEbIsResetting = 44,
    sizeof_ThunkData = 48,
    countof_pfn = 8
};
enum PushParamThunkData {
    sizeof_PushParamThunk = 20,
    countof_PushParamThunk = 256 // power of 2
};

__declspec(naked, noinline)
void __stdcall FireOnceTimerThunk(int hWnd, int uMsg, int wParam, int lParam)
{
    __asm {
__start:
        call    __next
__next:
        pop     edx
        sub     edx, 5
        sub     edx, offset __start
        push    edi
        push    esi

        // init thunk data
        mov     esi, dword ptr [esp+20]         // param wParam
        add     esi, 8
        cmp     dword ptr [esi], 0              // wParam->pfnCoTaskMemAlloc
        jz      __skip_init_thunk_data
        mov     edi, edx
        add     edi, offset __vtable
        mov     eax, edx
        add     eax, offset __QI
        stosd                                   // vtbl->pfnQI
        mov     eax, edx
        add     eax, offset __AddRef
        stosd                                   // vtbl->pfnAddRef
        mov     eax, edx
        add     eax, offset __Release
        stosd                                   // vtbl->pfnRelease
        xor     eax, eax
        stosd                                   // vtbl->PushParamIdx
        mov     ecx, countof_pfn
        rep     movsd                           // vtbl->pfn*
__skip_init_thunk_data:

        // this = CoTaskMemAlloc(sizeof_InstanceData)
        add     edx, offset __vtable
        push    edx
        push    sizeof_InstanceData
        call    dword ptr [edx+t_pfnCoTaskMemAlloc]
        pop     edx
        mov     edi, eax                        // edi = this ptr

        // init instance members
        mov     eax, edx
        stosd                                   // this->pVtbl
        mov     eax, 1
        stosd                                   // this->dwRefCnt
        xor     eax, eax
        stosd                                   // this->dwTimerID
        stosd                                   // this->pfnTimerProc
        mov     esi, dword ptr [esp+20]         // param wParam
        movsd                                   // wParam->pCallbackThis
        movsd                                   // wParam->pfnCallback
        sub     edi, sizeof_InstanceData        // rewind this ptr

        // init pfnPushParamThunk
        mov     eax, dword ptr [edx+t_PushParamIdx]
        dec     eax
        and     eax, countof_PushParamThunk - 1
        push    eax
__loop_find_empty_slot:
        mov     ecx, dword ptr [edx+t_PushParamIdx]
        cmp     ecx, dword ptr [esp]
        jne     __has_empty_slot
        // no empty slot available
        pop     eax
        mov     ecx, dword ptr [edi]            // ecx = this->pVtbl
        push    edi
        call    dword ptr [ecx+t_pfnCoTaskMemFree]
        xor     edi, edi                        // *lParam = NULL
        jmp     __exit_init
__has_empty_slot:
        inc     dword ptr [edx+t_PushParamIdx]
        and     dword ptr [edx+t_PushParamIdx], countof_PushParamThunk - 1
        lea     eax, [edx + 8*ecx]
        lea     eax, [eax + 8*ecx]
        lea     ecx, [eax + 4*ecx + sizeof_ThunkData] // ecx = pfnPushParamThunk
        cmp     byte ptr [ecx+sizeof_PushParamThunk-1], 0
        jne     __loop_find_empty_slot
        pop     eax
        mov     dword ptr [ecx+0], 0xB82434FF
        mov     dword ptr [ecx+4], edi
        mov     dword ptr [ecx+8], 0x04244489
        mov     eax, edx
        sub     eax, offset __vtable
        add     eax, offset __TimerProc
        push    eax
        shl     eax, 8
        add     eax, 0x000000B8
        mov     dword ptr [ecx+12], eax
        pop     eax
        shr     eax, 24
        add     eax, 0x90E0FF00
        mov     dword ptr [ecx+16], eax
        mov     dword ptr [edi+m_pfnPushParamThunk], ecx

        // this->dwTimerID = SetTimer(0, 0, Delay, this->pfnPushParamThunk)
        push    ecx                             // for lpTimerFunc
        push    dword ptr [esp+20]              // param wMsg (for Delay)
        push    0                               // for nIDEvent
        push    0                               // for hWnd
        mov     ecx, dword ptr [edi]            // ecx = this->pVtbl
        call    dword ptr [ecx+t_pfnSetTimer]
        mov     dword ptr [edi+m_dwTimerID], eax

        // retval & epilog
__exit_init:
        mov     eax, dword ptr [esp+24]         // param lParam
        mov     dword ptr [eax], edi            // *lParam = this
        pop     esi
        pop     edi
        mov     eax, offset __PushParamThunk
        sub     eax, offset __start
        add     eax, countof_PushParamThunk * sizeof_PushParamThunk
        ret     16

        align   4
__QI:
        mov     eax, dword ptr [esp+8]          // param riid
        cmp     dword ptr [eax+0], 0
        jne     __nointerface
        cmp     dword ptr [eax+4], 0
        jne     __nointerface
        cmp     dword ptr [eax+8], 0xC0
        jne     __nointerface
        cmp     dword ptr [eax+12], 0x46000000
        jne     __nointerface
        mov     edx, dword ptr [esp+4]          // param this
        inc     dword ptr [edx+m_dwRefCnt]
        mov     eax, dword ptr [esp+12]         // param ppvObject
        mov     dword ptr [eax], edx            // *ppvObject = this
        xor     eax, eax
        ret     12
__nointerface:
        mov     eax, 0x80004002                 // E_NOINTERFACE
        ret     12

        align   4
__AddRef:
        mov     edx, dword ptr [esp+4]          // param this
        inc     dword ptr [edx+m_dwRefCnt]
        mov     eax, dword ptr [edx+m_dwRefCnt]
        ret     4

        align   4
__Release:
        mov     edx, dword ptr [esp+4]          // param this
        dec     dword ptr [edx+m_dwRefCnt]
        mov     eax, dword ptr [edx+m_dwRefCnt]
        jnz     __exit_Release

        // clear PushParamThunk slot
        mov     eax, dword ptr [edx+m_pfnPushParamThunk]
        mov     byte ptr [eax+sizeof_PushParamThunk-1], 0
        
        // invoke KillTimer(0, this->dwTimerId)
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    dword ptr [edx+m_dwTimerID]     // this->TimerID (for uIDEvent)
        push    0                               // 0 (for hWnd)
        call    dword ptr [ecx+t_pfnKillTimer]
        
        // invoke CoTaskMemFree(this)
        mov     edx, dword ptr [esp+4]          // param this
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    edx
        call    dword ptr [ecx+t_pfnCoTaskMemFree]
        xor     eax, eax
__exit_Release:
        ret     4

        align   4
__TimerProc:
        mov     edx, dword ptr [esp+4]          // this ptr
        // if EbMode() > 1 goto __skip_callback
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        mov     eax, dword ptr [ecx+t_pfnEbMode]
        test    eax, eax
        jz      __call_callback
        push    edx
        call    eax                             // this->pfnEbMode
        pop     edx
        cmp     eax, 1                          // 1 = running
        ja      __skip_callback
        // if EbIsResetting() goto __skip_callback
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    edx
        call    dword ptr [ecx+t_pfnEbIsResetting]
        pop     edx
        test    eax, eax
        jnz      __skip_callback
        // invoke GetWindowLong(this->hIdeOwner, GWL_STYLE)
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    edx
        push    GWL_STYLE
        push    dword ptr [ecx+t_hIdeOwner]
        call    dword ptr [ecx+t_pfnGetWindowLong]
        pop     edx
        test    eax, WS_DISABLED
        jnz     __skip_callback
__call_callback:
        // invoke KillTimer(0, this->dwTimerId)
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    edx
        push    dword ptr [edx+m_dwTimerID]     // this->dwTimerID (for uIDEvent)
        push    0                               // 0 (for hWnd)
        call    dword ptr [ecx+t_pfnKillTimer]
        pop     edx
        inc     dword ptr [edx+m_dwRefCnt]      // __AddRef
        xor     eax, eax                        // pRetVal
        push    eax
        push    esp
        push    dword ptr [edx+m_pCallbackThis]
        call    dword ptr [edx+m_pfnCallback]
        mov     edx, dword ptr [esp+8]          // restore this ptr
        mov     dword ptr [edx+m_dwTimerID], 0  // note: clear this->TimerId *after* calling this->pfnCallback
        push    edx                             //       this keeps instance alive past this->pfnCallback call
        call    __Release
        pop     eax                             // RetVal
__skip_callback:
        ret     20
        align   4
__vtable:
        lea     eax,[eax+eax*2+1]               // t_pfnQI
        lea     eax,[eax+eax*2+1]               // t_pfnAddRef
        lea     eax,[eax+eax*2+1]               // t_pfnRelease
        lea     eax,[eax+eax*2+1]               // t_PushParamIdx
        lea     eax,[eax+eax*2+1]               // t_pfnCoTaskMemAlloc
        lea     eax,[eax+eax*2+1]               // t_pfnCoTaskMemFree
        lea     eax,[eax+eax*2+1]               // t_pfnSetTimer
        lea     eax,[eax+eax*2+1]               // t_pfnKillTimer
        lea     eax,[eax+eax*2+1]               // t_hIdeOwner
        lea     eax,[eax+eax*2+1]               // t_pfnGetWindowLong
        lea     eax,[eax+eax*2+1]               // t_pfnEbMode
        lea     eax,[eax+eax*2+1]               // t_pfnEbIsResetting
__PushParamThunk:
        //push    dword ptr [esp]
        //mov     eax, 0xDEADBEEF
        //mov     dword ptr [esp+4], eax
        //mov     eax, 0xDEADBEEF
        //jmp     eax
    }
}

#define THUNK_SIZE ((char *)main - (char *)FireOnceTimerThunk)

HRESULT __stdcall TimerProc(void *self, LRESULT *pRetVal);

void main()
{
    CoInitialize(0);
    void *vtbl[] = { TimerProc };
    void *pObj[] = { vtbl };
    InParams p = { 0 };
    p.pCallbackThis = pObj;
    p.pfnCallback = TimerProc;
    p.pfnCoTaskMemAlloc = GetProcAddress(GetModuleHandle(L"ole32"), "CoTaskMemAlloc");
    p.pfnCoTaskMemFree = GetProcAddress(GetModuleHandle(L"ole32"), "CoTaskMemFree");
    p.pfnSetTimer = GetProcAddress(GetModuleHandle(L"user32"), "SetTimer");
    p.pfnKillTimer = GetProcAddress(GetModuleHandle(L"user32"), "KillTimer");
    p.hIdeOwner = 0x1C520D8C;
    p.pfnGetWindowLong = GetProcAddress(GetModuleHandle(L"user32"), "GetWindowLongA");
    p.pfnEbMode = EbMode;
    p.pfnEbIsResetting = EbIsResetting;
    void *hThunk = VirtualAlloc(0, 0x10000, MEM_COMMIT, PAGE_EXECUTE_READWRITE);
    printf("hThunk=%p\nTHUNK_SIZE=%d\n", hThunk, THUNK_SIZE);
    memcpy(hThunk, FireOnceTimerThunk, THUNK_SIZE);
    IUnknown *pUnk = 0;
    DWORD dwSize = CallWindowProc((WNDPROC)hThunk, 0, 0xABC, (WPARAM)&p, (LPARAM)&pUnk);
    printf("dwSize=%d\nsizeof_InstanceData=%d\n", dwSize, sizeof_InstanceData);
    dwSize -= countof_PushParamThunk * sizeof_PushParamThunk;
    printf("offset __PushParamThunk=%d\n", dwSize);
    dwSize -= sizeof_ThunkData;
    printf("offset __vtable=%d\n", dwSize);
    MSG uMsg;
    while (GetMessage(&uMsg, 0, 0, 0) != 0) {
        TranslateMessage(&uMsg);
        DispatchMessage(&uMsg);
        // note: TimerID is cleared *after* calling pfnCallback so it's accessible
        //       through ThunkPrivateDate inside VB's TimerProc impl
        //if (*(int *)((BYTE *)pUnk + m_dwTimerID) == 0)
            break;
    }
    printf("pUnk->AddRef=%d\n", pUnk->AddRef());
    printf("pUnk->Release=%d\n", pUnk->Release());
    printf("pUnk->Release=%d\n", pUnk->Release());
    WCHAR szBuffer[1000];
    DWORD dwBufSize = _countof(szBuffer);
    CryptBinaryToString((BYTE *)FireOnceTimerThunk, dwSize, CRYPT_STRING_BASE64, szBuffer, &dwBufSize);
    for(int i = 0, j = 0; (szBuffer[j] = szBuffer[i]) != 0; )
        ++i, j += (szBuffer[j] != '\r' && szBuffer[j] != '\n');
    printf("Const STR_THUNK     As String = \"%S\" ' %S\n", szBuffer, GetCurrentDateTime());
}

HRESULT __stdcall TimerProc(void *self, LRESULT *pRetVal)
{
    printf("TimerProc\n");
    *pRetVal = 123;
    return S_OK;
}
#endif

#ifdef IMPL_CLEANUP_THUNK
struct InParams {
    void *pfnCoTaskMemAlloc;
    void *pfnCoTaskMemFree;
};
enum InstanceData {
    m_pVtbl = 0,
    m_dwRefCnt = 4,
    m_hHandle = 8,
    m_pfnCleanup = 12,
    sizeof_InstanceData = 16,
};
enum ThunkData {
    t_pfnQI = 0,
    t_pfnAddRef = 4,
    t_pfnRelease = 8,
    t_pfnCoTaskMemAlloc = 12,
    t_pfnCoTaskMemFree = 16,
    sizeof_ThunkData = 20,
    countof_pfn = 2
};

__declspec(naked, noinline)
void __stdcall CleanupThunk(int hWnd, int uMsg, int wParam, int lParam)
{
    __asm {
__start:
        call    __next
__next:
        pop     edx
        sub     edx, 5
        sub     edx, offset __start
        push    edi
        push    esi

        // init thunk data
        mov     esi, dword ptr [esp+20]         // param wParam
        cmp     dword ptr [esi], 0              // wParam->pfnCoTaskMemAlloc
        jz      __skip_init_thunk_data
        mov     edi, edx
        add     edi, offset __vtable
        mov     eax, edx
        add     eax, offset __QI
        stosd                                   // vtbl->pfnQI
        mov     eax, edx
        add     eax, offset __AddRef
        stosd                                   // vtbl->pfnAddRef
        mov     eax, edx
        add     eax, offset __Release
        stosd                                   // vtbl->pfnRelease
        mov     ecx, countof_pfn
        rep     movsd                           // vtbl->pfn*
__skip_init_thunk_data:

        // this = CoTaskMemAlloc(sizeof_InstanceData)
        add     edx, offset __vtable
        push    edx
        push    sizeof_InstanceData
        call    dword ptr [edx+t_pfnCoTaskMemAlloc]
        pop     edx
        mov     edi, eax                        // edi = this ptr

        // init instance members
        mov     eax, edx
        stosd                                   // this->pVtbl
        mov     eax, 1
        stosd                                   // this->dwRefCnt
        mov     eax, dword ptr [esp+12]         // param hWnd
        stosd                                   // this->hHandle
        mov     eax, dword ptr [esp+16]         // param uMsg
        stosd                                   // this->pfnCleanup
        sub     edi, sizeof_InstanceData        // rewind this ptr

        // retval & epilog
        mov     eax, dword ptr [esp+24]         // param lParam
        mov     dword ptr [eax], edi            // *lParam = this
        pop     esi
        pop     edi
        mov     eax, offset __end
        sub     eax, offset __start
        ret     16

        align   4
__QI:
        mov     eax, dword ptr [esp+8]          // eax = param riid
        cmp     dword ptr [eax+0], 0
        jne     __nointerface
        cmp     dword ptr [eax+4], 0
        jne     __nointerface
        cmp     dword ptr [eax+8], 0xC0
        jne     __nointerface
        cmp     dword ptr [eax+12], 0x46000000
        jne     __nointerface
        mov     edx, dword ptr [esp+4]          // edx = param this
        inc     dword ptr [edx+m_dwRefCnt]
        mov     eax, dword ptr [esp+12]         // eax = param ppvObject
        mov     dword ptr [eax], edx            // *ppvObject = this
        xor     eax, eax
        ret     12
__nointerface:
        mov     eax, 0x80004002                 // E_NOINTERFACE
        ret     12

        align   4
__AddRef:
        mov     edx, dword ptr [esp+4]          // edx = param this
        inc     dword ptr [edx+m_dwRefCnt]
        mov     eax, dword ptr [edx+m_dwRefCnt]
        ret     4

        align   4
__Release:
        mov     edx, dword ptr [esp+4]          // edx = param this
        dec     dword ptr [edx+m_dwRefCnt]
        mov     eax, dword ptr [edx+m_dwRefCnt]
        jnz     __exit_Release
        
        // invoke this->pfnCleanup(this->hHandle)
        push    dword ptr [edx+m_hHandle]
        call    dword ptr [edx+m_pfnCleanup]
        
        // invoke CoTaskMemFree(this)
        mov     edx, dword ptr [esp+4]          // edx = param this
        mov     ecx, dword ptr [edx]            // ecx = this->pVtbl
        push    edx
        call    dword ptr [ecx+t_pfnCoTaskMemFree] // vtbl->pfnCoTaskMemFree
        xor     eax, eax
__exit_Release:
        ret     4

        align   4
__vtable:
        lea     eax,[eax+eax*2+1]               // t_pfnQI
        lea     eax,[eax+eax*2+1]               // t_pfnAddRef
        lea     eax,[eax+eax*2+1]               // t_pfnRelease
        lea     eax,[eax+eax*2+1]               // t_pfnCoTaskMemAlloc
        lea     eax,[eax+eax*2+1]               // t_pfnCoTaskMemFree
__end:
    }
}

#define THUNK_SIZE ((char *)main - (char *)CleanupThunk)

void main()
{
    CoInitialize(0);
    void *vtbl[] = { 0 };
    void *pObj[] = { vtbl };
    InParams p = { 0 };
    p.pfnCoTaskMemAlloc = GetProcAddress(GetModuleHandle(L"ole32"), "CoTaskMemAlloc");
    p.pfnCoTaskMemFree = GetProcAddress(GetModuleHandle(L"ole32"), "CoTaskMemFree");
    HWND hWnd = CreateWindowEx(0, L"STATIC", L"Test", 0, 0, 0, 0, 0, 0, 0, 0, 0);
    printf("hWnd=%p\n", hWnd);
    ShowWindow(hWnd, SW_SHOWMAXIMIZED);
    void *hThunk = VirtualAlloc(0, 0x10000, MEM_COMMIT, PAGE_EXECUTE_READWRITE);
    printf("hThunk=%p\nTHUNK_SIZE=%d\n", hThunk, THUNK_SIZE);
    memcpy(hThunk, CleanupThunk, THUNK_SIZE);
    IUnknown *pUnk = 0;    
    DWORD dwSize = CallWindowProc((WNDPROC)hThunk, hWnd, (UINT)GetProcAddress(GetModuleHandle(L"user32"), "DestroyWindow"), (WPARAM)&p, (LPARAM)&pUnk);
    printf("dwSize=%d\nsizeof_InstanceData=%d\n", dwSize, sizeof_InstanceData);
    printf("pUnk->AddRef=%d\n", pUnk->AddRef());
    printf("pUnk->Release=%d\n", pUnk->Release());
    printf("pUnk->Release=%d\n", pUnk->Release());
    WCHAR szBuffer[1000];
    DWORD dwBufSize = _countof(szBuffer);
    CryptBinaryToString((BYTE *)CleanupThunk, dwSize - sizeof_ThunkData, CRYPT_STRING_BASE64, szBuffer, &dwBufSize);
    for(int i = 0, j = 0; (szBuffer[j] = szBuffer[i]) != 0; )
        ++i, j += (szBuffer[j] != '\r' && szBuffer[j] != '\n');
    printf("Const STR_THUNK     As String = \"%S\" ' %S\n", szBuffer, GetCurrentDateTime());
}
#endif
