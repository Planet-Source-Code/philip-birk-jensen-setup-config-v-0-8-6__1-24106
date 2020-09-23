VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fAnalyse 
   Caption         =   "Analyse"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   Icon            =   "analyse.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   5280
   Begin MSComctlLib.ImageList imgSmall 
      Left            =   3330
      Top             =   1620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "analyse.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "analyse.frx":02A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "analyse.frx":03FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmAnalyse 
      Height          =   2625
      Left            =   0
      TabIndex        =   2
      Top             =   -45
      Width           =   5280
      Begin MSComctlLib.ListView lvwAnalyse 
         Height          =   2445
         Left            =   45
         TabIndex        =   3
         Top             =   135
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   4313
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgSmall"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Level"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Information"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.CommandButton cmdAnalyse 
      Caption         =   "Analyse"
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Top             =   2610
      Width           =   5190
   End
   Begin MSComctlLib.ProgressBar pbrStatus 
      Height          =   150
      Left            =   0
      TabIndex        =   1
      Top             =   2925
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "fAnalyse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnalyse_Click()
   pbrStatus.Value = 0
   StartAnalyse lvwAnalyse
End Sub

Private Sub Form_Resize()

   frmAnalyse.Width = fAnalyse.ScaleWidth
   frmAnalyse.Height = fAnalyse.ScaleHeight - 460

   lvwAnalyse.Width = fAnalyse.ScaleWidth - 100
   lvwAnalyse.Height = frmAnalyse.Height - 200
   
   cmdAnalyse.Width = fAnalyse.ScaleWidth - 85
   cmdAnalyse.Top = frmAnalyse.Height
   
   pbrStatus.Width = fAnalyse.ScaleWidth
   pbrStatus.Top = frmAnalyse.Height + 300

End Sub
