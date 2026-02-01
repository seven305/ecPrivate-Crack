Attribute VB_Name = "ApFace"
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, _
    ByVal wBar As Long, ByVal bShow As Long) As Long
    
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_CHILD = &H40000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_SYSMENU = &H80000
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const WM_PASTE = &H302

Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

Private Const SB_HORZ = 0
Private Const SB_VERT = 1
Private Const SB_BOTH = 3
 
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_STATICEDGE = &H20000

Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

Private Const BS_FLAT = &H8000&

Global Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Const WM_USER = &H400
Const CCM_FIRST = &H2000&
Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Const PBM_SETBARCOLOR = (WM_USER + 9)

Public Function Change_pb_Color(ByVal hwnd As Long, _
                                ByVal lColor As Long)
          SendMessage hwnd, PBM_SETBKCOLOR, 0, ByVal lColor
End Function
Public Function Change_pb_ForeColor(ByVal hwnd As Long, _
                                    ByVal lColor As Long)
          SendMessage hwnd, PBM_SETBARCOLOR, 0, ByVal lColor
End Function
Public Sub MakeFlat(lhWnd As Long)
    Dim lStyle As Long
    
    ' Get window style
    lStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
    ' Setup window styles
    lStyle = lStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    ' Set window style
    SetWindowLong lhWnd, GWL_EXSTYLE, lStyle
    RemoveBorder lhWnd
End Sub
Public Sub RemoveBorder(lhWnd As Long)
    Dim lStyle As Long
    
    ' Get window style
    lStyle = GetWindowLong(lhWnd, GWL_STYLE)
    ' Setup window styles
    lStyle = lStyle And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION Or WS_BORDER Or WS_SIZEBOX Or WS_THICKFRAME)
    ' Set window style
    SetWindowLong lhWnd, GWL_STYLE, lStyle
    ' Update window
    SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub

Function btnFlat(Button As CommandButton)
    SetWindowLong Button.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
    Button.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function


Sub ListBoxVertScroll(lbListBox As ListBox, bShow As Boolean)
    
    If bShow Then
        Call ShowScrollBar(lbListBox.hwnd, SB_VERT, 1&)
    Else
        Call ShowScrollBar(lbListBox.hwnd, SB_VERT, 0&)
    End If
End Sub


