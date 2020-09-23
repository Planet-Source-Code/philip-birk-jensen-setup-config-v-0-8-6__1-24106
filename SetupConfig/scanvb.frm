VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form ScanVB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Scan Tool"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   3705
   Begin VB.Frame frmInfo 
      Height          =   2805
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   3705
      Begin MSComctlLib.ListView lvwProject 
         Height          =   2535
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Information"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "ScanVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This will analyse the VB project file.
'
' I will add some more feutures in the future, maybe, not sure, because
' I only created this tool for fun.
'
' To be honest I don't think that I will keep on expanding this tool
' it's to boring, it's just typing and typing and typing.
'
' Later on it might be able to scan the single objects, in
' the project, and give you some info there.
'

Private Type ProjectItems
   pForms     As Integer
   pModules   As Integer
   pClass     As Integer
   pMajor     As Integer
   pMinor     As Integer
   pRevision  As Integer
End Type

Dim pItems As ProjectItems


Private Sub Form_Load()
Dim sTemp As String

   Open sVBProject For Input As #1
      Do Until EOF(1) = True
         Line Input #1, sTemp
         Select Case CutLeft(sTemp)
            Case "Type"
               lvwProject.ListItems.Add , "Typ", "Type"
               lvwProject.ListItems("Typ").ListSubItems.Add , , CutRight(sTemp)
            Case "Title"
               lvwProject.ListItems.Add , "Tit", "Title"
               lvwProject.ListItems("Tit").ListSubItems.Add , , CutRight(sTemp)
            Case "VersionComments"
               lvwProject.ListItems.Add , "Com", "Comments"
               lvwProject.ListItems("Com").ListSubItems.Add , , CutRight(sTemp)
            Case "Form"
               pItems.pForms = pItems.pForms + 1
            Case "Module"
               pItems.pModules = pItems.pModules + 1
            Case "Class"
               pItems.pClass = pItems.pClass + 1
         End Select
      Loop
   Close #1
   
   If pItems.pForms <> 0 Then
      lvwProject.ListItems.Add , "For", "Forms"
      lvwProject.ListItems("For").ListSubItems.Add , , pItems.pForms
   End If
   If pItems.pModules <> 0 Then
      lvwProject.ListItems.Add , "Mod", "Modules"
      lvwProject.ListItems("Mod").ListSubItems.Add , , pItems.pModules
   End If
   If pItems.pClass <> 0 Then
      lvwProject.ListItems.Add , "Cla", "Classes"
      lvwProject.ListItems("Cla").ListSubItems.Add , , pItems.pClass
   End If
   
End Sub
