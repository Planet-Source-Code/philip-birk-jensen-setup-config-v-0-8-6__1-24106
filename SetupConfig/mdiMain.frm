VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm mdiMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Setup Config"
   ClientHeight    =   7350
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8760
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sbrInfo 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   7110
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2461
            MinWidth        =   2469
            Object.ToolTipText     =   "File path"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1323
            MinWidth        =   1323
            Object.ToolTipText     =   "Size"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "Created"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   4095
      Top             =   3375
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0A24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0DBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1158
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":14F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":188C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1C26
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMenu 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   510
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   4
            Object.Width           =   400
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   400
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "&Select"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Text            =   "&Open"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Text            =   "&Compile"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Text            =   "&Run"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Text            =   "&Locate Dir"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picIcon 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      Picture         =   "mdiMain.frx":1D80
      ScaleHeight     =   450
      ScaleWidth      =   8700
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8760
   End
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   945
      Top             =   2070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open setup.lst"
      Filter          =   "setup.lst|*.lst|All Files|*.*"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mFile 
         Caption         =   "&Open"
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mFile 
         Caption         =   "&Save"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mFile 
         Caption         =   "Op&tions"
         Index           =   4
      End
      Begin VB.Menu mFile 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mFile 
         Caption         =   "&Exit"
         Index           =   6
      End
   End
   Begin VB.Menu mnuVB 
      Caption         =   "&Visual Basic"
      Begin VB.Menu mVB 
         Caption         =   "&Select Project"
         Index           =   0
      End
      Begin VB.Menu mVB 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mVB 
         Caption         =   "&Open Project"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mVB 
         Caption         =   "&Compile Project"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mVB 
         Caption         =   "&Run Project"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mVB 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mVB 
         Caption         =   "&Locate Dir"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu mVB 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mVB 
         Caption         =   "&Scan Utility"
         Enabled         =   0   'False
         Index           =   8
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
      Begin VB.Menu mProject 
         Caption         =   "setup.lst &Data"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mProject 
         Caption         =   "setup.lst &Explore"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mProject 
         Caption         =   "setup.lst &Clean"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mProject 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mProject 
         Caption         =   "&Tools"
         Enabled         =   0   'False
         Index           =   4
         Begin VB.Menu mTools 
            Caption         =   "&Analyse"
            Index           =   0
         End
         Begin VB.Menu mTools 
            Caption         =   "&Date utility"
            Index           =   1
         End
      End
      Begin VB.Menu mProject 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mProject 
         Caption         =   "&Test Setup"
         Enabled         =   0   'False
         Index           =   6
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mWindow 
         Caption         =   "Tile &Horizontally"
         Index           =   0
      End
      Begin VB.Menu mWindow 
         Caption         =   "Tile &Vertically"
         Index           =   1
      End
      Begin VB.Menu mWindow 
         Caption         =   "&Cascade"
         Index           =   2
      End
      Begin VB.Menu mWindow 
         Caption         =   "&Arrange Icons"
         Index           =   3
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mHelp 
         Caption         =   "Contents"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mHelp 
         Caption         =   "About"
         Index           =   1
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuList 
         Caption         =   "List"
         Begin VB.Menu mList 
            Caption         =   "View"
            Index           =   0
            Begin VB.Menu mView 
               Caption         =   "Large Icons"
               Index           =   0
            End
            Begin VB.Menu mView 
               Caption         =   "Small Icons"
               Index           =   1
            End
            Begin VB.Menu mView 
               Caption         =   "List"
               Index           =   2
            End
            Begin VB.Menu mView 
               Caption         =   "Details"
               Index           =   3
            End
         End
         Begin VB.Menu mList 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mList 
            Caption         =   "Locate in Clean"
            Index           =   2
         End
      End
      Begin VB.Menu mnuTree 
         Caption         =   "Tree"
         Begin VB.Menu mTree 
            Caption         =   "Locate in Clean"
            Index           =   0
         End
         Begin VB.Menu mTree 
            Caption         =   "-"
            Index           =   1
         End
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ################################################################################
'   Author: Philip Birk-Jensen
'   E-Mail: vtg@sol.dk
'
'   I made this program because I didn't feel, that I had enough control over
'   my setup files. And I don't see any reason to get another setup program,
'   when there is one following with VB. Ok well I do see a lot of other
'   reasons, but I'm just weird, so I still use the one with VB.
'
'   So if you thought you just downloed a new setup program, then you'r wrong
'   This program just tweaks/edits your current setup programs, made with the
'   utility following VB.
'
'   Even though this can edit the standard setup, it's not all editing. It have
'   to follow the frames made by the utility it self. This meaning that you
'   can't set in pictures, start program when done, and other very useful stuff
'   just run it and see what it does.
'
' ------------------------------------------------------------------------------
'
'   So far it only edits in existing project. But I expect it to make it's own
'   sometime (still under restrictions of standard VB utility).
'
' ###############################################################################


Option Explicit

Private Sub LoadProject()
On Error GoTo CancelError:
   cdlFile.Filter = "Visual Basic Project|*.vbp"
   cdlFile.DialogTitle = "Open Visual Basic Project"
   cdlFile.ShowOpen
   sVBProject = cdlFile.FileName
   cdlFile.Filter = "setup.lst|*.lst|All Files|*.*"
   cdlFile.DialogTitle = "Open setup.lst"
   FileLoaded 1
CancelError:
End Sub

Private Sub MDIForm_Load()

   mdiMain.Icon = picIcon.Picture
   If GetSetting("SetupConfig", "StartUp", "ShowTip", 1) = 1 Then
      Tip.Show 0, Me
   End If
   
   sFile = "C:\Documents and Settings\Philip\Desktop\Vb\PSCode\Setup Setup\Setup\setup.lst"
 
   DateU.Show
 
   cdlFile.Filter = "setup.lst|*.lst|All Files|*.*"
   cdlFile.DialogTitle = "Open setup.lst"
   
   sVBPath = GetSetting("SetupConfig", "Options", "VBPath", "")
   sVBEXE = GetSetting("SetupConfig", "Options", "VBEXE", "")
   
End Sub

Private Sub mHelp_Click(Index As Integer)
   Select Case Index
      Case 0
      Case 1
         About.Show 1, Me
   End Select
End Sub

Private Sub mList_Click(Index As Integer)
On Error GoTo Err1
Dim iTmp As Integer
Dim iFile As Integer
   Select Case Index
      Case 2
         iTmp = InStr(1, Clean.txtData.Text, Tree.lvwDetails.SelectedItem.Text) - 1
         Clean.txtData.SelStart = iTmp
         Clean.txtData.SelLength = Len(Tree.lvwDetails.SelectedItem.Text)
         Clean.SetFocus
   End Select
Err1:
   If Err.Number <> 0 Then
      Clean.Show
      Resume
   End If
End Sub

Private Sub mTools_Click(Index As Integer)
   Select Case Index
      Case 0 ' Analyse
         fAnalyse.Show
      Case 1
         DateU.Show
   End Select
End Sub

Private Sub mVB_Click(Index As Integer)
On Error GoTo ErrorHandler
   Select Case Index
      Case 0 ' Select Project
         LoadProject
      Case 2
         VBOpen
      Case 3
         VBCompile
      Case 4
         VBRunP
      Case 6
         VBDir
   End Select
ErrorHandler:
   If Err.Number = 32755 Then
   ElseIf Err.Number <> 0 Then
      Debug.Print ; "++ERROR++"
      Debug.Print "  Number: "; Err.Number
      Debug.Print "  Description: "; Err.Description
   End If
End Sub

Private Sub mView_Click(Index As Integer)
   
   Select Case Index
      Case 0
         Tree.lvwDetails.View = 0
      Case 1
         Tree.lvwDetails.View = 1
      Case 2
         Tree.lvwDetails.View = 2
      Case 3
         Tree.lvwDetails.View = 3
   End Select
   
End Sub

Private Sub mFile_Click(Index As Integer)
On Error GoTo ErrorHandler
Dim iFile As Integer
   
   Select Case Index
      Case 0 ' Open
         cdlFile.ShowOpen
         sFile = cdlFile.FileName & "0"
         FileCopy cdlFile.FileName, sFile
         FileLoaded 0
      Case 2 ' Save
         iFile = FreeFile
         Open sFile For Output As iFile
            Print iFile, Clean.txtData
         Close iFile
      Case 4 ' options
         Options.Show 1, Me
   End Select
   
ErrorHandler:
   If Err.Number = 32755 Then
   ElseIf Err.Number <> 0 Then
      Debug.Print ; "++ERROR++"
      Debug.Print "  Number: "; Err.Number
      Debug.Print "  Description: "; Err.Description
   End If
End Sub

Private Sub mProject_Click(Index As Integer)
Dim sTemp As String

   Select Case Index
      Case 0 ' Data
         Data.Show
      Case 1 ' Tree
         Tree.Show
      Case 2 ' Clean
         Clean.Show
      Case 6 ' Test Setup
         sTemp = Replace(LCase(sFile), "setup.lst0", "")
         Shell sTemp & "setup.exe", vbNormalFocus
   End Select
End Sub

Private Sub mWindow_Click(Index As Integer)
   
   Select Case Index
      Case 0 ' Horizontally
         mdiMain.Arrange 1
      Case 1 ' Vertically
         mdiMain.Arrange 2
      Case 2 ' Cascade
         mdiMain.Arrange 0
      Case 3 ' Arrange Icons
         mdiMain.Arrange 3
   End Select
   
End Sub

Private Sub tbrMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo CancelError:
   Select Case Button.Index
      Case 2 ' Open
         cdlFile.ShowOpen
         sFile = cdlFile.FileName & "0"
         FileCopy cdlFile.FileName, sFile
         FileLoaded 0
      Case 5 ' Data
         Data.Show
      Case 6 ' Tree
         Tree.Show
      Case 7 ' clean
         Clean.Show
      Case 8 ' Analyse
         fAnalyse.Show
      Case 10 ' VB
         LoadProject
   End Select

CancelError:
End Sub

Private Sub tbrMenu_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

   Select Case ButtonMenu.Index
      Case 1
         LoadProject
      Case 3
         VBOpen
      Case 4
         VBCompile
      Case 5
         VBRunP
      Case 7
         VBDir
   End Select
   
End Sub
