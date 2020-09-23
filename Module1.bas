Attribute VB_Name = "Module1"
Option Explicit

Public Enum IDM
    a = 128
    b
    c
    d
    e
End Enum
Public procOld As Long
Private Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, ByVal hWnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&)
Private Const WM_SYSCOMMAND = &H112

Public Function WindowProc(ByVal hWnd As Long, _
ByVal iMsg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long

   Select Case iMsg
      Case WM_SYSCOMMAND
         Select Case wParam
         Case IDM.a
            MsgBox "'First menu' Cliked"
         Case IDM.b
            MsgBox "'Option Before Close' Cliked"
         Case IDM.c
            MsgBox "GoodBye"
            Unload Form1
            Exit Function
         'Case IDM.d: MsgBox "Clik en 'Option MaRiO' Cliked" 'Disabled...
         Case IDM.e
            MsgBox "'Last Option' Cliked"
         End Select
   End Select

    WindowProc = CallWindowProc(procOld, hWnd, iMsg, wParam, lParam)
End Function




