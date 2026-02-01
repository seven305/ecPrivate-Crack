Attribute VB_Name = "ModExtra"
'|| AOL 9.0 API Module Coded by: Seven
'||
'|| If you are going to use this for a program, Please give me sexy
'|| Credits :)
'|| Esoteric Code
'|| http://www.sevenz.net
'|| <seven@sevenz.net>
Public Declare Function GetMenu Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "User32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "User32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SendMessageByNum& Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessageByString Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLong& Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function GetWindowTextLength Lib "User32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindow Lib "User32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function EnableWindow Lib "User32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long

Public Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Public Declare Function FlashWindow Lib "User32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Public Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long


Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_SETTEXT = &HC
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_GETITEMDATA = &H199

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5

Public Const VK_SPACE = &H20
Public Const VK_DOWN = &H28
Public Const VK_RETURN = &HD

Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200

Public Const WM_SYSCOMMAND = &H112
Public Const WM_USER = &H400

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
'===================
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

Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_EX_STATICEDGE = &H20000
'===================
' Public constants for SetWindowPos API declaration
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

' Public constants for ShowWindow API declaration
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_RESTORE = 9
'===================
Public Const AOL_IGNORE = &H9
Public Const AOL_EJECT = &HA

Public RICHCNTL As Long
Public AOLChild As Long
Public MDIClient As Long
Public aolframe As Long
Public AOLFrame25 As Long
Public AOLIcon As Long
Public InternetExplorerServer As Long
Public ATLFD As Long
Public ATLFDD As Long
Public AOLIndWnd As Long
Public ATLBA As Long

Public ATL69F381D8 As Long
Public RICHCNTLREADONLY As Long
Public AOLFontCombo As Long

Public AOLListbox As Long
Public strBuffer As String
Public lngRetVal As Long
Public lngLength As Long

Public AteClass As Long
Public WndAteClass As Long
Public AIMIMessage As Long
Public MMUIWnd As Long
Public OscarIconBtn As Long
Public OscarTree As Long

Public AIMChatWnd As Long


Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Public Enum SoundTypes
    SoundSync = &H0
    SoundASync = &H1
    SoundMemory = &H4
    SoundLoop = &H8
    SoundNoStop = &H10
    SoundNoDefault = &H2
End Enum


Public Sub PlaySound()

    sndPlaySound StrConv(LoadResData(101, "CUSTOM"), vbUnicode), SoundASync Or SoundMemory
End Sub

'=========================

Public Sub AOL9_ChatSend(ByVal mMessage As String)

aolframe = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(aolframe, 0, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0, "AOL Child", vbNullString)
RICHCNTL = FindWindowEx(AOLChild, 0, "RICHCNTL", vbNullString)

aolframe = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(aolframe, 0, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0, "AOL Child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0, "_AOL_Icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)

Call SendMessageByString(RICHCNTL, WM_SETTEXT, 0&, mMessage)
Call SendMessage(RICHCNTL, WM_CHAR, VK_RETURN, 0)

ChatSend mMessage
End Sub
Public Function FindRoom() As Long

    Dim aolframe As Long, MDIClient As Long, AOLChild As Long
    aolframe = FindWindow("aol frame25", vbNullString)
    MDIClient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
    Dim Winkid1 As Long, Winkid2 As Long, Winkid3 As Long, Winkid4 As Long, Winkid5 As Long, Winkid6 As Long, Winkid7 As Long, Winkid8 As Long, Winkid9 As Long, FindOtherWin As Long
    FindOtherWin = GetWindow(AOLChild, GW_HWNDFIRST)
    Do While FindOtherWin <> 0
           DoEvents
           Winkid1 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
           Winkid2 = FindWindowEx(FindOtherWin, 0&, "richcntlreadonly", vbNullString)
           Winkid3 = FindWindowEx(FindOtherWin, 0&, "_aol_combobox", vbNullString)
           Winkid4 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
           Winkid5 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
           Winkid6 = FindWindowEx(FindOtherWin, 0&, "richcntl", vbNullString)
           Winkid7 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
           Winkid8 = FindWindowEx(FindOtherWin, 0&, "_aol_image", vbNullString)
           Winkid9 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
           If (Winkid1 <> 0) And (Winkid2 <> 0) And (Winkid3 <> 0) And (Winkid4 <> 0) And (Winkid5 <> 0) And (Winkid6 <> 0) And (Winkid7 <> 0) And (Winkid8 <> 0) And (Winkid9 <> 0) Then
                  FindRoom = FindOtherWin
                  Exit Function
           End If
           FindOtherWin = GetWindow(FindOtherWin, GW_HWNDNEXT)
    Loop
    FindRoom = 0
  
End Function
Public Sub AOL_GetChatList(ListToGet As Long, ListToAddTo As ListBox, Optional AddUser As Boolean = False)

    On Error Resume Next
    
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, Index As Long
    Dim sThread As Long, mThread As Long
    
    sThread& = GetWindowThreadProcessId(ListToGet, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0 To ListCount(ListToGet) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(ListToGet, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmHold& = itmHold& + 28
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            ScreenName$ = LCase(Replace(ScreenName$, " ", vbNullString))
            
            If ScreenName$ <> frmMain.txtNo And Len(ScreenName$) <> 0 Then
                ListToAddTo.AddItem ScreenName$
            End If
            
        Next Index&
        Call CloseHandle(mThread)
    End If
End Sub
Public Function GetUSer() As String

aolframe = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(aolframe, 0, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0, "AOL Child", vbNullString)

lngLength = GetWindowTextLength(AOLChild)
strBuffer = String(lngLength + 1, " ")
lngRetVal = SendMessageByString(AOLChild, WM_GETTEXT, lngLength + 1, strBuffer)
strBuffer = Left(strBuffer, lngLength)

 GetUSer = LCase(Mid(strBuffer, 10))

End Function

Public Sub ChatOptions(strUser As String, ByVal m_Command As Integer, Optional blnPartial As Boolean = False)

On Error Resume Next
    
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lngCheckBox As Long, lngChatUserInfo As Long
    
    strUser = LCase(strUser)

    rList = AOLList
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0 To ListCount(rList) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmHold& = itmHold& + 28
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = LCase(Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1))
            
            If (blnPartial = True And InStr(ScreenName, strUser)) Or (blnPartial = False And ScreenName = strUser) Then
                Call ListlClick(rList, CInt(Index&))
                
                
                aolframe = FindWindow("AOL Frame25", vbNullString)
                MDIClient = FindWindowEx(aolframe, 0, "MDIClient", vbNullString)
                AOLChild = FindWindowEx(MDIClient, 0, "AOL Child", vbNullString)
                AOLIcon = FindWindowEx(AOLChild, 0, "_AOL_Icon", vbNullString)
                
                  For i = 1 To m_Command
                     AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
                  Next i
                  
                  Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
                  Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
                
                
                Call CloseWin(lngChatUserInfo)
                Call CloseHandle(mThread)
                
                Exit Sub
            End If

        Next Index&
        Call CloseHandle(mThread)
    End If
    
End Sub

Private Sub CloseWin(hwnd As Long)
    Call SendMessageLong(hwnd, WM_CLOSE, 0&, 0&)
End Sub
Public Function AOLList() As Long
aolframe = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(aolframe, 0, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0, "AOL Child", vbNullString)
AOLList = FindWindowEx(AOLChild, 0, "_AOL_Listbox", vbNullString)

End Function
Private Function ListCount(ListBox As Long) As Long
    ListCount& = SendMessageLong(ListBox&, LB_GETCOUNT, 0&, 0&)
End Function
Public Sub AOL_CloseChilds()

aolframe = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(aolframe, 0, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0, "AOL Child", vbNullString)

Call SendMessageLong(AOLChild, WM_CLOSE, 0&, 0&)

End Sub

Public Function AOL9_KeyWord(ByVal m_Key As String)

aolframe = FindWindow("AOL Frame25", vbNullString)
AOLToolbar = FindWindowEx(aolframe, 0, "AOL Toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar, 0, "_AOL_Toolbar", vbNullString)
AOLEdit = FindWindowEx(AOLToolbar, 0, "_AOL_Edit", vbNullString)

Call SendMessageByString(AOLEdit, WM_SETTEXT, 0&, vbNullString)
Call SendMessageByString(AOLEdit, WM_SETTEXT, 0&, m_Key)

Call SendMessage(AOLEdit, WM_CHAR, VK_RETURN, 0)

End Function


Public Sub ChatSend(ByVal m_ChatMessage As String)
AIMChatWnd = FindWindow("AIM_ChatWnd", vbNullString)
WndAteClass = FindWindowEx(AIMChatWnd, 0, "WndAte32Class", vbNullString)
WndAteClass = FindWindowEx(AIMChatWnd, WndAteClass, "WndAte32Class", vbNullString)
AteClass = FindWindowEx(WndAteClass, 0, "Ate32Class", vbNullString)

Call SendMessageByString(AteClass, WM_SETTEXT, 0&, m_ChatMessage)
ChatButton
End Sub

Private Sub ChatButton()
AIMChatWnd = FindWindow("AIM_ChatWnd", vbNullString)
OscarIconBtn = FindWindowEx(AIMChatWnd, 0, "_Oscar_IconBtn", vbNullString)
OscarIconBtn = FindWindowEx(AIMChatWnd, OscarIconBtn, "_Oscar_IconBtn", vbNullString)
OscarIconBtn = FindWindowEx(AIMChatWnd, OscarIconBtn, "_Oscar_IconBtn", vbNullString)
OscarIconBtn = FindWindowEx(AIMChatWnd, OscarIconBtn, "_Oscar_IconBtn", vbNullString)
OscarIconBtn = FindWindowEx(AIMChatWnd, OscarIconBtn, "_Oscar_IconBtn", vbNullString)


Call PostMessage(OscarIconBtn, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(OscarIconBtn, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Function AOL9_GetLastChatLine() As String
On Error Resume Next
Dim strBuffer As String
Dim lngLength As Long
Dim mBuffer As String
Dim m_ChatLastLine As String

aolframe = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(aolframe, 0, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0, "AOL Child", vbNullString)
RICHCNTLREADONLY = FindWindowEx(AOLChild, 0, "RICHCNTLREADONLY", vbNullString)



    lngLength& = SendMessageLong(RICHCNTLREADONLY, WM_GETTEXTLENGTH, 0&, 0&)
    strBuffer$ = String(lngLength& + 1, Chr(0))
    Call SendMessageByString(RICHCNTLREADONLY, WM_GETTEXT, lngLength& + 1, strBuffer$)
    strBuffer$ = Left(strBuffer$, lngLength&)

    
       mBuffer = strBuffer
       m_ChatLastLine = Right(mBuffer, Len(mBuffer) - InStrRev(mBuffer, Chr(13)))
       AOL9_GetLastChatLine = m_ChatLastLine
       
End Function

Public Sub AOL9_OpenMail()
aolframe = FindWindow("AOL Frame25", vbNullString)
AOLToolbar = FindWindowEx(aolframe, 0, "AOL Toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar, 0, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, 0, "_AOL_Icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
    Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub

Private Sub ListlClick(lngList As Long, intIndex As Integer)
    Call SendMessageLong(lngList, LB_SETCURSEL, intIndex, 0&)
End Sub
Private Sub ListlDbClick(lngList As Long, intIndex As Integer)
    Call SendMessageLong(lngList, LB_SETCURSEL, intIndex, 0&)
    Call SendMessageLong(lngList, WM_LBUTTONDOWN, intIndex, 0&)
End Sub


Public Function AOL9_GetChatCaption() As String

aolframe = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(aolframe, 0, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0, "AOL Child", vbNullString)

lngLength = GetWindowTextLength(AOLChild)
strBuffer = String(lngLength + 1, " ")
lngRetVal = SendMessageByString(AOLChild, WM_GETTEXT, lngLength + 1, strBuffer)
strBuffer = Left(strBuffer, lngLength)

AOL9_GetChatCaption = strBuffer
End Function
Public Sub AOL9_Expression()
aolframe = FindWindow("AOL Frame25", vbNullString)
AOLToolbar = FindWindowEx(aolframe, 0, "AOL Toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar, 0, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, 0, "_AOL_Icon", vbNullString)
For i = 1 To 10
  AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
Next i

      Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
      Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Function AOL9_SendIM(ByVal m_UserName As String, ByVal m_Message As String, ByVal CloseWindow As Boolean) '
Dim lngButton As Integer

aolframe = FindWindow("AOL Frame25", vbNullString)
AOLToolbar = FindWindowEx(aolframe, 0, "AOL Toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar, 0, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, 0, "_AOL_Icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)

      Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
      Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
      
Delay 0.6
aolframe = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(aolframe, 0, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0, "AOL Child", vbNullString)
ATLBA = FindWindowEx(AOLChild, 0, "ATL:6724B2A0", vbNullString)
AOLIndWnd = FindWindowEx(ATLBA, 0, "_AOL_IndWnd", vbNullString)
AOLEdit = FindWindowEx(AOLIndWnd, 0, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit, WM_SETTEXT, 0&, m_UserName)


Delay 0.6
aolframe = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(aolframe, 0, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0, "AOL Child", vbNullString)
ATLBA = FindWindowEx(AOLChild, 0, "ATL:6724B2A0", vbNullString)
AOLIndWnd = FindWindowEx(ATLBA, 0, "_AOL_IndWnd", vbNullString)
RICHCNTL = FindWindowEx(AOLIndWnd, 0, "RICHCNTL", vbNullString)
Call SendMessageByString(RICHCNTL, WM_SETTEXT, 0&, m_Message)
Call SendMessage(RICHCNTL, WM_CHAR, VK_RETURN, 0)


Delay 0.4

aolframe = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(aolframe, 0, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0, "AOL Child", vbNullString)
ATLBA = FindWindowEx(AOLChild, 0, "ATL:6724B2A0", vbNullString)
AOLIndWnd = FindWindowEx(ATLBA, 0, "_AOL_IndWnd", vbNullString)
AOLIcon = FindWindowEx(AOLIndWnd, 0, "_AOL_Icon", vbNullString)

For lngButton = 1 To 18
   AOLIcon = FindWindowEx(AOLIndWnd, AOLIcon, "_AOL_Icon", vbNullString)
Next lngButton

      Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
      Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)

Delay 0.6
DoEvents
 If CloseWindow = True Then
   Call SendMessageLong(AOLChild, WM_CLOSE, 0&, 0&)
 End If
 
End Function

Sub OpenURL(lol)
ShellExecute hwnd, "open", lol, vbNullString, vbNullString, SW_SHOWMAXIMIZED
End Sub



Public Function RGB2Hex(r As Byte, g As Byte, b As Byte) As String
    On Error Resume Next
    ' convert to long using vb's rgb function, then use the long2rgb function
    RGB2Hex = Long2Hex(RGB(r, g, b))
End Function


Public Function Long2Hex(LongColor As Long) As String
    On Error Resume Next
    ' use vb's hex function
    Long2Hex = Hex(LongColor)
End Function


Public Sub Delay(ByVal Float As Single)
If Float > 10 Then Stop
Float = Timer + Float
While Float >= Timer
  DoEvents
Wend
End Sub

Sub Pause(interval)
    Current = Timer
    Do While Timer - Current < Val(interval)
    DoEvents
    Loop
End Sub
