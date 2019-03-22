VERSION 5.00
Begin VB.Form frmTestTimers 
   Caption         =   "Form2"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form2"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2604
      Top             =   1428
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   936
      Left            =   924
      TabIndex        =   0
      Top             =   588
      Width           =   1356
   End
End
Attribute VB_Name = "frmTestTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_pTimer As IUnknown

Private Property Get AddressOfTimerProc() As frmTestTimers
    Set AddressOfTimerProc = InitAddressOfMethod(Me, 0)
End Property

Private Sub Command1_Click()
    Debug.Print "here"
    Set m_pTimer = InitFireOnceTimerThunk(ObjPtr(Me), AddressOfTimerProc.TimerProc(), Delay:=1000)
End Sub

Private Sub Form_Load()
'    Dim o As Object
    Set m_pTimer = InitFireOnceTimerThunk(ObjPtr(Me), AddressOfTimerProc.TimerProc(), Delay:=100)
'    Set o = m_pTimer
End Sub

Public Function TimerProc() As Long
    Debug.Print "TimerProc, App.NonModalAllowed=" & App.NonModalAllowed, Timer
'    Set m_pTimer = InitFireOnceTimerThunk(ObjPtr(Me), AddressOfTimerProc.TimerProc(), Delay:=100)
    Set m_pTimer = Nothing
End Function

Private Sub Timer1_Timer()
    Debug.Print "Timer1_Timer", Timer
'    Timer1.Enabled = False
End Sub
