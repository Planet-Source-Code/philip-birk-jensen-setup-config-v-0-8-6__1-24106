VERSION 5.00
Begin VB.Form Data 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup Config"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "data.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5205
   WhatsThisHelp   =   -1  'True
   Begin VB.ListBox lstDir 
      Height          =   1230
      ItemData        =   "data.frx":038A
      Left            =   2925
      List            =   "data.frx":03A0
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   2130
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
      Left            =   4995
      TabIndex        =   17
      Top             =   2610
      Width           =   195
   End
   Begin VB.Frame frmEditors 
      Caption         =   "Editors"
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5190
      Begin VB.CommandButton cmdGroup 
         Caption         =   "Group Editor"
         Height          =   330
         Left            =   3780
         TabIndex        =   16
         Top             =   270
         Width           =   1230
      End
      Begin VB.CommandButton cmdFiles 
         Caption         =   "Setup Files"
         Height          =   330
         Left            =   2565
         TabIndex        =   15
         Top             =   270
         Width           =   1230
      End
      Begin VB.CommandButton cmdBootFiles 
         Caption         =   "Boot Files"
         Height          =   330
         Left            =   1350
         TabIndex        =   4
         Top             =   270
         Width           =   1230
      End
      Begin VB.CommandButton cmdBoot 
         Caption         =   "Boot Section"
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   270
         Width           =   1230
      End
   End
   Begin VB.Frame frmSetup 
      Caption         =   "Setup Settings"
      Height          =   1770
      Left            =   0
      TabIndex        =   0
      Top             =   810
      Width           =   5190
      Begin VB.TextBox txtUninstall 
         Height          =   285
         Left            =   1215
         TabIndex        =   13
         Top             =   1395
         Width           =   3885
      End
      Begin VB.TextBox txtExe 
         Height          =   285
         Left            =   1215
         TabIndex        =   11
         Top             =   1080
         Width           =   3885
      End
      Begin VB.CheckBox chkForce 
         Caption         =   "Use the default dir (user will not be asked)"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   855
         Width           =   5010
      End
      Begin VB.CommandButton cmdDir 
         Height          =   195
         Left            =   4950
         TabIndex        =   8
         Top             =   585
         Width           =   150
      End
      Begin VB.TextBox txtDir 
         Height          =   285
         Left            =   1215
         TabIndex        =   6
         Top             =   540
         Width           =   3660
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   1215
         TabIndex        =   1
         Top             =   225
         Width           =   3885
      End
      Begin VB.Label lblUninstall 
         Caption         =   "Uninstall Text:"
         Height          =   240
         Left            =   90
         TabIndex        =   14
         Top             =   1440
         Width           =   5010
      End
      Begin VB.Label lblExe 
         Caption         =   "App Exe:"
         Height          =   240
         Left            =   90
         TabIndex        =   12
         Top             =   1125
         Width           =   5010
      End
      Begin VB.Label lblDir 
         Caption         =   "Default Dir:"
         Height          =   240
         Left            =   90
         TabIndex        =   7
         Top             =   585
         Width           =   5010
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title:"
         Height          =   240
         Left            =   90
         TabIndex        =   2
         Top             =   270
         Width           =   5010
      End
   End
End
Attribute VB_Name = "Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sLength(3) As String

Private Sub chkForce_Click()
Dim iTmp As Integer
   
   If InStr(1, Clean.txtData.Text, "ForceUseDefDir") <> 0 Then
      UpdateClean "ForceUseDefDir", "1", chkForce
   Else
      iTmp = InStr(1, Clean.txtData.Text, "AppToUninstall=")
      Clean.txtData.SelStart = iTmp + 15 + Len(sLength(3))
      Clean.txtData.SelText = Chr(10) & "ForceUseDefDir=1" & Chr(13)
   End If
         
End Sub

Private Sub cmdBoot_Click()

   BootStrap.Show

End Sub

Private Sub cmdBootFiles_Click()

   sFiles = "Bootstrap Files"
   Files.Show

End Sub

Private Sub cmdDir_Click()

   lstDir.Visible = True
   lstDir.SetFocus

End Sub

Private Sub cmdFiles_Click()

   sFiles = "Setup1 Files"
   Files.Show

End Sub

Private Sub Form_Load()
Dim sLine As String
Dim sTmp As String
Dim iFile As Integer

   iFile = FreeFile
   
   Open sFile For Input As iFile
      Do Until sLine = "[Setup]"
      
         Line Input #iFile, sLine
      
      Loop
      
      Line Input #iFile, sLine
      If CutLeft(sLine) = "ForceUseDefDir" Then
         ForceIt (sLine)
         Line Input #iFile, sLine
         sTmp = CutRight(sLine)
         txtTitle.Text = sTmp
         sLength(0) = sTmp
      Else
         sTmp = CutRight(sLine)
         txtTitle.Text = sTmp
         sLength(0) = sTmp
      End If
      
      Line Input #iFile, sLine
      If CutLeft(sLine) = "ForceUseDefDir" Then
         ForceIt (sLine)
         Line Input #iFile, sLine
         sTmp = CutRight(sLine)
         txtDir.Text = sTmp
         sLength(1) = sTmp
      Else
      sTmp = CutRight(sLine)
      txtDir.Text = sTmp
      sLength(1) = sTmp
      End If
      
      Line Input #iFile, sLine
      If CutLeft(sLine) = "ForceUseDefDir" Then
         ForceIt (sLine)
         Line Input #iFile, sLine
         sTmp = CutRight(sLine)
         txtExe.Text = sTmp
         sLength(2) = sTmp
      Else
      sTmp = CutRight(sLine)
      txtExe.Text = sTmp
      sLength(2) = sTmp
      End If
      
      Line Input #iFile, sLine
      If CutLeft(sLine) = "ForceUseDefDir" Then
         ForceIt (sLine)
         Line Input #iFile, sLine
         sTmp = CutRight(sLine)
         txtUninstall.Text = sTmp
         sLength(3) = sTmp
      Else
      sTmp = CutRight(sLine)
      txtUninstall.Text = sTmp
      sLength(3) = sTmp
      End If
      
      Line Input #iFile, sLine
      If CutLeft(sLine) = "ForceUseDefDir" Then
         ForceIt (sLine)
      End If
      
   Close iFile
End Sub

Private Sub ForceIt(sLine As String)
On Error Resume Next
Dim sTmp As String
   sTmp = CutRight(sLine)
   chkForce.Value = sTmp
End Sub

Private Sub lstDir_DblClick()

   txtDir.Text = lstDir.List(lstDir.ListIndex) & "\" & txtTitle.Text
   lstDir.Visible = False
   txtDir.SetFocus

End Sub

Private Sub lstDir_KeyPress(KeyAscii As Integer)

   If KeyAscii = 13 Then
      lstDir_DblClick
   End If

End Sub

Private Sub txtDir_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      UpdateTitle "DefaultDir", sLength(1), txtDir.Text
      sLength(1) = txtDir.Text
   End If
End Sub

Private Sub txtDir_LostFocus()
   UpdateTitle "DefaultDir", sLength(1), txtDir.Text
   sLength(1) = txtDir.Text
End Sub

Private Sub txtExe_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      UpdateTitle "AppExe", sLength(2), txtExe.Text
      sLength(2) = txtExe.Text
   End If

End Sub

Private Sub txtExe_LostFocus()
   UpdateTitle "AppExe", sLength(2), txtExe.Text
   sLength(2) = txtExe.Text
End Sub

Private Sub txtTitle_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = 13 Then
      UpdateTitle "Title", sLength(0), txtTitle.Text
      sLength(0) = txtTitle.Text
   End If
   
End Sub

Private Sub txtTitle_LostFocus()
      UpdateTitle "Title", sLength(0), txtTitle.Text
      sLength(0) = txtTitle.Text
End Sub

Private Sub txtUninstall_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      UpdateTitle "AppToUninstall", sLength(3), txtUninstall.Text
      sLength(3) = txtUninstall.Text
   End If
End Sub

Private Sub txtUninstall_LostFocus()
   UpdateTitle "AppToUninstall", sLength(3), txtUninstall.Text
   sLength(3) = txtUninstall.Text
End Sub
