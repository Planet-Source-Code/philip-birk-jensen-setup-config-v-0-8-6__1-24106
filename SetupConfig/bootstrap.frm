VERSION 5.00
Begin VB.Form Bootstrap 
   BorderStyle     =   0  'None
   Caption         =   "Boot Section"
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2445
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmButtons 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   6045
      Begin VB.CommandButton cmdMini 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5580
         TabIndex        =   6
         Top             =   135
         Width           =   195
      End
      Begin VB.CommandButton cmdSpawn 
         Caption         =   "Processing Cab file and Uninstall ->>"
         Height          =   285
         Left            =   1260
         TabIndex        =   5
         Top             =   1350
         Width           =   3435
      End
      Begin VB.CommandButton cmdDone 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5805
         TabIndex        =   4
         Top             =   135
         Width           =   195
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5760
         TabIndex        =   3
         Top             =   2205
         Width           =   195
      End
      Begin VB.TextBox txtSetupText 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   195
         Left            =   1260
         TabIndex        =   2
         Top             =   1080
         Width           =   3435
      End
      Begin VB.TextBox txtSetupTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   195
         Left            =   45
         TabIndex        =   1
         Top             =   135
         Width           =   5955
      End
   End
End
Attribute VB_Name = "Bootstrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Const HTCAPTION = 2
Dim sLength(1) As String

Private Sub cmdDone_Click()

   Unload Me

End Sub

Private Sub cmdMini_Click()

   Bootstrap.WindowState = 1

End Sub

Private Sub cmdSpawn_Click()

   Bootafter.Show 0, mdiMain

End Sub

Private Sub Form_Load()
Dim sLine As String
Dim sTmp As String
Dim ID As Integer
Dim iFile As Integer

   iFile = FreeFile
   
   Open sFile For Input As iFile
      Do Until sLine = "[Bootstrap]"
      
         Line Input #iFile, sLine
      
      Loop
      
      Line Input #iFile, sLine
      sTmp = CutRight(sLine)
      txtSetupTitle.Text = sTmp
      sLength(0) = sTmp
      
      Line Input #iFile, sLine
      sTmp = CutRight(sLine)
      txtSetupText.Text = sTmp
      sLength(1) = sTmp
      
   Close iFile

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ReleaseCapture
   Dim DragWindow

   DragWindow = SendMessage(Bootstrap.hwnd, &HA1, HTCAPTION, 0&) ' Just had this
   'from and old progam, so don't know about the &HA1 and 0&.

End Sub

Private Sub frmButtons_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   Form_MouseDown Button, Shift, x, y

End Sub

Private Sub txtSetupText_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = 13 Then
      UpdateClean "SetupText", sLength(1), txtSetupText.Text
      sLength(1) = txtSetupText.Text
   End If
   
End Sub

Private Sub txtSetupText_LostFocus()

   UpdateClean "SetupText", sLength(1), txtSetupText.Text
   sLength(1) = txtSetupText.Text

End Sub

Private Sub txtSetupTitle_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = 13 Then
      UpdateClean "SetupTitle", sLength(0), txtSetupTitle.Text
      sLength(0) = txtSetupTitle.Text
   End If
   
End Sub

Private Sub txtSetupTitle_LostFocus()

      UpdateClean "SetupTitle", sLength(0), txtSetupTitle.Text
      sLength(0) = txtSetupTitle.Text

End Sub
