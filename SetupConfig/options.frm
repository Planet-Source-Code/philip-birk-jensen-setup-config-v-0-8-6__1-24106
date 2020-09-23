VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Options 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkShowTip 
      Caption         =   "Show tips at startup"
      Height          =   195
      Left            =   225
      TabIndex        =   11
      Top             =   585
      Width           =   2805
   End
   Begin VB.Frame frmOther 
      Caption         =   "Other Options"
      Height          =   1905
      Left            =   90
      TabIndex        =   10
      ToolTipText     =   "Visual Basic options."
      Top             =   315
      Width           =   4470
   End
   Begin VB.Frame frmVB 
      Caption         =   "Visual Basic"
      Height          =   1905
      Left            =   90
      TabIndex        =   5
      ToolTipText     =   "Visual Basic options."
      Top             =   315
      Width           =   4470
      Begin VB.CommandButton cmdDetect 
         Caption         =   "Detect"
         Height          =   275
         Left            =   2835
         TabIndex        =   9
         ToolTipText     =   "Try to auto detect VB. It might not be found."
         Top             =   810
         Width           =   735
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   275
         Left            =   3645
         TabIndex        =   8
         ToolTipText     =   "Search after the directory"
         Top             =   810
         Width           =   735
      End
      Begin VB.TextBox txtVBPath 
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Text            =   "Visual Basic Directory"
         ToolTipText     =   "VB Path"
         Top             =   495
         Width           =   4290
      End
      Begin VB.Label lblVBPath 
         Caption         =   "Visual Basic path/directory"
         Height          =   240
         Left            =   90
         TabIndex        =   7
         ToolTipText     =   "VB Path"
         Top             =   225
         Width           =   4290
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   330
      Left            =   3960
      TabIndex        =   1
      ToolTipText     =   "Activate configurations, but don't quiet the Options box."
      Top             =   2385
      Width           =   690
   End
   Begin MSComctlLib.TabStrip tabOptions 
      Height          =   2310
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   4075
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Visual Basic"
            Object.ToolTipText     =   "Setup all the information needed for the program to support VB."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Other Configurations"
            Object.ToolTipText     =   "All other configurations, this program dosn't have much to configure."
            ImageVarType    =   2
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3285
      TabIndex        =   2
      ToolTipText     =   "Quiet without applying the changes."
      Top             =   2385
      Width           =   690
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   330
      Left            =   2610
      TabIndex        =   3
      ToolTipText     =   "Quiet and use the current settings."
      Top             =   2385
      Width           =   690
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Information box"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   2340
      Width           =   2580
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkShowTip_Click()
   If chkShowTip.Value = 0 Then
      SaveSetting "SetupConfig", "StartUp", "ShowTip", "0"
   Else
      SaveSetting "SetupConfig", "StartUp", "ShowTip", "1"
   End If
End Sub

Private Sub cmdApply_Click()
On Error Resume Next

   sVBEXE = Replace(txtVBPath.Text, """%1""", "")
   
   sVBPath = Replace(txtVBPath.Text, "vb6.exe"" ""%1", "")
   sVBPath = Replace(sVBPath, """", "")
   
   SaveSetting "SetupConfig", "Options", "VBPath", sVBPath
   SaveSetting "SetupConfig", "Options", "VBEXE", sVBEXE

End Sub

Private Sub cmdBrowse_Click()
Dim sTemp As String

   sTemp = BDirectory(Options)
   
   txtVBPath.Text = sTemp
   
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdDetect_Click()
Dim sTemp As String
   sTemp = GetString(HKEY_CLASSES_ROOT, "applications\vb6.exe\shell\open\command", "")

   If sTemp = "" Then
      sTemp = GetString(HKEY_CURRENT_USER, "applications\vb5.exe\shell\open\command", "")
   End If

   If sTemp = "" Then
      MsgBox "Unable to auto detect the VB directory, please search manualy", vbInformation + vbOKOnly, "Auto Detect VB"
   End If
   
   txtVBPath.Text = sTemp
   
End Sub

Private Sub cmdDone_Click()
   cmdApply_Click
   Unload Me
End Sub

Private Sub Form_Load()
   txtVBPath.Text = GetSetting("SetupConfig", "Options", "VBPath", "")
   chkShowTip.Value = GetSetting("SetupConfig", "StartUp", "ShowTip", 1)
End Sub

Private Sub tabOptions_Click()
   Select Case tabOptions.TabIndex
      Case 1
         frmVB.ZOrder 1
      Case 2
         frmOther.ZOrder 1
   End Select
End Sub
