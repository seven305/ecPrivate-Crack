VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   3090
   ClientLeft      =   7080
   ClientTop       =   3825
   ClientWidth     =   3375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   3375
   Begin VB.TextBox tB 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox tG 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox tR 
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   4200
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.CheckBox ChkSettings 
         Caption         =   "Play sound on new cracks"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   2175
      End
      Begin VB.ComboBox CmbTime 
         Height          =   315
         ItemData        =   "frmSettings.frx":0000
         Left            =   1440
         List            =   "frmSettings.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1560
         Width           =   615
      End
      Begin VB.CheckBox ChkSettings 
         Caption         =   "Auto Load cracks"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   2400
         Width           =   855
      End
      Begin VB.CheckBox ChkSettings 
         Caption         =   "Enable Chat sends"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton cmdDef 
         Caption         =   "Default"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   855
      End
      Begin MSComDlg.CommonDialog Cdgl 
         Left            =   2400
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtColor 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Text            =   "#C0C0C0"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00B07D4E&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   3
         Top             =   1920
         Width           =   375
      End
      Begin VB.CheckBox ChkSettings 
         Caption         =   "Append Mode on"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox ChkSettings 
         Caption         =   "Notify on new cracks"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Min(s)"
         Height          =   255
         Left            =   2160
         TabIndex        =   14
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Send Rates every"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Hex Code:"
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   2040
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdDef_Click()
picColor.BackColor = &HB07D4E
txtColor.Text = "#C0C0C0"
End Sub

Private Sub cmdSave_Click()


 Open App.Path & "\settings" For Output As #1
 
    For i = 0 To ChkSettings.UBound
     Print #1, ChkSettings(i).Value
    Next i
 Close #1
 
 Call SaveSetting("cracker", "settings", "color", txtColor.Text)
 Call SaveSetting("cracker", "settings", "time", CmbTime.ListIndex)
 Me.Hide
 
End Sub

Private Sub Form_Load()
MakeFlat cmdSave.hwnd
MakeFlat cmdDef.hwnd
MakeFlat txtColor.hwnd

 Dim success%
 
 success% = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
Me.Hide
End Sub

Private Sub picColor_Click()
 Cdgl.DialogTitle = "ecPrivate Cracker Settings"
 Cdgl.ShowColor
 
 picColor.BackColor = Cdgl.Color
 Dim RGBValue As Long
 Dim tmpColor As String
 RGBValue = Cdgl.Color
 
    tR.Text = RGBValue And &HFF&
    tG.Text = (RGBValue And &HFF00&) / 256
    tB.Text = (RGBValue And &HFF0000) / 65536
    'FF00
    tmpColor = Hex(RGB(tR, tG, tB))
    txtColor.Text = "#" & Hex(RGB(tR, tG, tB))
    
    If tmpColor = "FF" Then txtColor.Text = "#FF0000"
    
    
    
End Sub
