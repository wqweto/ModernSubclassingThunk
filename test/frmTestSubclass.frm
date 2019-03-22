VERSION 5.00
Begin VB.Form frmTestSubclass 
   Caption         =   "Form2"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form2"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   348
      Left            =   336
      TabIndex        =   1
      Top             =   420
      Width           =   2448
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   360
      Left            =   336
      TabIndex        =   0
      Top             =   1008
      UseMnemonic     =   0   'False
      Width           =   2424
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmTestSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WM_WINDOWPOSCHANGED As Long = &H47

Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private m_pSubclass1 As IUnknown
Private m_pSubclass2 As IUnknown

Public Function SubclassProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Attribute SubclassProc.VB_MemberFlags = "40"
    If wMsg = WM_WINDOWPOSCHANGED Then
        Label1.Caption = "wMsg=" & wMsg & "  " & Timer
    End If
    SubclassProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
End Function

Public Function frWndProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If wMsg = WM_WINDOWPOSCHANGED Then
        Label2.Caption = "wMsg=" & wMsg & "  " & Timer
    End If
    frWndProc = DefSubclassProc(hWnd, wMsg, wParam, lParam)
End Function

Private Property Get pvAddressOfSubclassProc() As frmTestSubclass
    Set pvAddressOfSubclassProc = InitAddressOfMethod(Me, 4)
End Property

Private Sub Form_Load()
'    Dim o As Object
    Set m_pSubclass2 = InitSubclassingThunk(hWnd, Me, AddressOf RedirectFrmTestSubclassWndProc)
    Set m_pSubclass1 = InitSubclassingThunk(hWnd, Me, pvAddressOfSubclassProc.SubclassProc(0, 0, 0, 0))
'    Set m_pSubclass2 = InitSubclassingThunk(hWnd, Me, pvAddressOfSubclassProc.frWndProc(0, 0, 0, 0))
'    Set o = m_pSubclass1
End Sub

Private Sub Form_Click()
    Set m_pSubclass1 = Nothing
End Sub
