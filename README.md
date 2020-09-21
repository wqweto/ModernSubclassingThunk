## The Modern Subclassing Thunk (MST)

IDE-safe subclassing, hooking and fire-once timers that have no problems with `End` button/statement in VBIDE.

### Description

These thunks allow implementing IDE-safe window subclassing, windows hooks and fire-once timers for self-contained user-controls, forms and classes.

### Usage

Start by including `src\mdModernSubclassing.bas` in your project.

#### Subclassing guidelines

Add private `m_pSubclass As IUnknown` member variable to your form/class.

Add private `pvAddressOfSubclassProc` helper property to your form/class.

```
Private Property Get pvAddressOfSubclassProc() As frmTestSubclass
    Set pvAddressOfSubclassProc = InitAddressOfMethod(Me, 5)
End Property
```

Notice return type `frmTestSubclass` is the name of your form/class and this is important to match.

Add public `SubclassProc` function with proper [SUBCLASSPROC](https://docs.microsoft.com/en-us/windows/win32/api/commctrl/nc-commctrl-subclassproc) input parameters and return type.

```
Public Function SubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
    Select Case wMsg
    Case WM_WINDOWPOSCHANGED
        ...
    End Select
End Function
```

Start subclassing with an `InitSubclassingThunk` call in `Form_Load` or any other routine.

```
    Set m_pSubclass = InitSubclassingThunk(hWnd, Me, pvAddressOfSubclassProc.SubclassProc(0, 0, 0, 0, 0))
```

Can explicitly unsubclass by clearing the `m_pSubclass` variable in `Form_Unload` but no harm to leave it auto-clear on shutdown.

```
    Set m_pSubclass = Nothing
```
