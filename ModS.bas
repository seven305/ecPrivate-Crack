Attribute VB_Name = "ModS"
'|||||||||||||| AIM

Private Sub FindAIMButton()
MMUIWnd = FindWindow("M:MUIWnd", vbNullString)
AIMIMessage = FindWindowEx(MMUIWnd, 0, "AIM_IMessage", vbNullString)
OscarIconBtn = FindWindowEx(AIMIMessage, 0, "_Oscar_IconBtn", vbNullString)

Call PostMessage(OscarIconBtn, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(OscarIconBtn, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub SendAIM_IM(ByVal m_Message As String)
AIMIMessage = FindWindowEx(MMUIWnd, 0, "AIM_IMessage", vbNullString)
WndAteClass = FindWindowEx(AIMIMessage, 0, "WndAte32Class", vbNullString)
WndAteClass = FindWindowEx(AIMIMessage, WndAteClass, "WndAte32Class", vbNullString)
AteClass = FindWindowEx(WndAteClass, 0, "Ate32Class", vbNullString)

Call SendMessageByString(AteClass, WM_SETTEXT, 0&, m_Message)
FindAIMButton
End Sub



Public Sub RunMenu(Main_Prog As String, Top_Position As String, Menu_String As String)


    On Error GoTo stp
    Dim Top_Position_Num As Long, buffer As String, Look_For_Menu_String As Long
    Dim Trim_Buffer As String, Sub_Menu_Handle As Long, BY_POSITION As Long, Get_ID As Long
    Dim Click_Menu_Item As Long, Menu_Parent As Long, AOL As Long, Menu_Handle As Long, Parent As Long
    

    Top_Position_Num = -1
    Parent& = FindWindow(Main_Prog, vbNullString)
    Menu_Handle = GetMenu(Parent&)
    Do
        DoEvents
        Top_Position_Num = Top_Position_Num + 1
        buffer$ = String$(255, 0)
        Look_For_Menu_String& = GetMenuString(Menu_Handle, Top_Position_Num, buffer$, Len(Top_Position) + 1, WM_USER)
        Trim_Buffer = FixAPIString(buffer$)
        If Trim_Buffer = Top_Position Then Exit Do
        If GetMenuItemID(Menu_Handle, Top_Position_Num) = 0 Then Exit Do
    Loop

    Sub_Menu_Handle = GetSubMenu(Menu_Handle, Top_Position_Num)
    BY_POSITION = -1
    Do
        DoEvents
        BY_POSITION = BY_POSITION + 1
        buffer$ = String(255, 0)
        Look_For_Menu_String& = GetMenuString(Sub_Menu_Handle, BY_POSITION, buffer$, Len(Menu_String) + 1, WM_USER)
        Trim_Buffer = FixAPIString(buffer$)
        If Trim_Buffer = Menu_String Then Exit Do
        If GetMenuItemID(Menu_Handle, BY_POSITION) = 0 Then Exit Do
    Loop
    DoEvents
    Get_ID& = GetMenuItemID(Sub_Menu_Handle, BY_POSITION)
    Click_Menu_Item = SendMessageByNum(Parent&, WM_COMMAND, Get_ID&, 0&)
stp:
End Sub
Public Function FixAPIString(strToFix As String) As String

     FixAPIString = Replace(strToFix, Chr(0), "")
End Function

