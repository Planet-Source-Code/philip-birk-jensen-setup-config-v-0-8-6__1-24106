VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Bootafter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BootStrap"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmTemp 
      Caption         =   "Location of temporary files"
      Height          =   915
      Left            =   0
      TabIndex        =   11
      Top             =   2655
      Width           =   4425
      Begin VB.CommandButton cmdTemp 
         Height          =   195
         Left            =   4185
         TabIndex        =   13
         Top             =   540
         Width           =   150
      End
      Begin VB.TextBox txtTemp 
         Height          =   285
         Left            =   135
         TabIndex        =   12
         Top             =   495
         Width           =   3975
      End
      Begin VB.Label lblTemp 
         Caption         =   "File to run after bootstrap finished"
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   270
         Width           =   4110
      End
   End
   Begin VB.Frame frmAfter 
      Caption         =   "Run after pre install"
      Height          =   915
      Left            =   0
      TabIndex        =   7
      Top             =   1710
      Width           =   4425
      Begin VB.TextBox txtSpawn 
         Height          =   285
         Left            =   135
         TabIndex        =   9
         Top             =   495
         Width           =   3975
      End
      Begin VB.CommandButton cmdSpawn 
         Height          =   195
         Left            =   4185
         TabIndex        =   8
         Top             =   540
         Width           =   150
      End
      Begin VB.Label lblAfter 
         Caption         =   "File to run after bootstrap finished"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   270
         Width           =   4110
      End
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
      Left            =   4230
      TabIndex        =   4
      Top             =   3600
      Width           =   195
   End
   Begin VB.Frame frmCab 
      Caption         =   "Cab File(s)"
      Height          =   1680
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4425
      Begin MSComctlLib.Slider sldCabs 
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Top             =   1080
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   503
         _Version        =   393216
         LargeChange     =   4
         Min             =   1
         Max             =   21
         SelStart        =   1
         TickFrequency   =   2
         Value           =   1
      End
      Begin VB.CommandButton cmdPath 
         Height          =   195
         Left            =   4185
         TabIndex        =   3
         Top             =   540
         Width           =   150
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   135
         TabIndex        =   1
         Top             =   495
         Width           =   3975
      End
      Begin VB.Label lbl20 
         Caption         =   "21"
         Height          =   195
         Left            =   4095
         TabIndex        =   16
         Top             =   1395
         Width           =   195
      End
      Begin VB.Label lbl0 
         Caption         =   "0"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   1395
         Width           =   150
      End
      Begin VB.Label lblAmount 
         Caption         =   "Amount of Cab files"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   855
         Width           =   4110
      End
      Begin VB.Label lblPath 
         Caption         =   "Cab file"
         Height          =   195
         Left            =   135
         TabIndex        =   2
         Top             =   270
         Width           =   4110
      End
   End
   Begin MSComDlg.CommonDialog cdlLocate 
      Left            =   2880
      Top             =   2295
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "Bootafter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sLength(3) As String

Private Sub cmdPath_Click()
On Error GoTo CancelEnd

   cdlLocate.DialogTitle = "Locate the cab file"
   cdlLocate.Filter = "Cab File|*.cab|All Files|*.*"
   cdlLocate.ShowOpen
   txtPath.Text = cdlLocate.FileTitle
   txtPath_LostFocus
   
CancelEnd:
End Sub

Private Sub cmdSpawn_Click()
On Error GoTo CancelEnd

   cdlLocate.DialogTitle = "File to run after pre install"
   cdlLocate.Filter = "Application|*.exe|WinHelp File|*.hlp|HTMLHelp File|*.chm|Text File|*.txt|All Files|*.*"
   cdlLocate.ShowOpen
   txtSpawn.Text = cdlLocate.FileTitle
   txtSpawn_LostFocus
   
CancelEnd:
End Sub

Private Sub cmdTemp_Click()

   txtTemp.Text = BDirectory(Bootafter)
   txtTemp_LostFocus

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
      Line Input #iFile, sLine
      
      Line Input #iFile, sLine
      sTmp = CutRight(sLine)
      txtPath.Text = sTmp
      sLength(0) = sTmp
      
      Line Input #iFile, sLine
      sTmp = CutRight(sLine)
      txtSpawn.Text = sTmp
      sLength(1) = sTmp
      
      Line Input #iFile, sLine
      
      Line Input #iFile, sLine
      sTmp = CutRight(sLine)
      txtTemp.Text = sTmp
      sLength(2) = sTmp
      
      Line Input #iFile, sLine
      sTmp = CutRight(sLine)
      sldCabs.Value = sTmp
      sLength(3) = sTmp
      
   Close iFile

End Sub


Private Sub sldCabs_KeyUp(KeyCode As Integer, Shift As Integer)

   UpdateClean "Cabs", sLength(3), sldCabs.Value
   sLength(3) = sldCabs.Value

End Sub

Private Sub sldCabs_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   UpdateClean "Cabs", sLength(3), sldCabs.Value
   sLength(3) = sldCabs.Value

End Sub

Private Sub txtPath_KeyUp(KeyCode As Integer, Shift As Integer)

   If KeyCode = 13 Then
      UpdateClean "CabFile", sLength(0), txtPath.Text
      sLength(0) = txtPath.Text
   End If

End Sub

Private Sub txtPath_LostFocus()
   
   UpdateClean "CabFile", sLength(0), txtPath.Text
   sLength(0) = txtPath.Text

End Sub

Private Sub txtSpawn_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = 13 Then
      UpdateClean "Spawn", sLength(1), txtSpawn.Text
      sLength(1) = txtSpawn.Text
   End If
   
End Sub

Private Sub txtSpawn_LostFocus()

   UpdateClean "Spawn", sLength(1), txtSpawn.Text
   sLength(1) = txtSpawn.Text

End Sub

Private Sub txtTemp_KeyUp(KeyCode As Integer, Shift As Integer)

   If KeyCode = 13 Then
      UpdateClean "TmpDir", sLength(2), txtTemp.Text
      sLength(2) = txtTemp.Text
   End If

End Sub

Private Sub txtTemp_LostFocus()

   UpdateClean "TmpDir", sLength(2), txtTemp.Text
   sLength(2) = txtTemp.Text

End Sub
