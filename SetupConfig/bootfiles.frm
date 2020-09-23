VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Files 
   Caption         =   "Files"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   Icon            =   "bootfiles.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4215
   ScaleWidth      =   6015
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   240
      Left            =   4320
      TabIndex        =   10
      Top             =   3465
      Width           =   1230
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
      Left            =   5580
      TabIndex        =   9
      Top             =   3465
      Width           =   195
   End
   Begin VB.Frame frmBoot 
      Height          =   3435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5775
      Begin VB.CheckBox chkShared 
         Caption         =   "Shared"
         Height          =   285
         Left            =   45
         TabIndex        =   8
         Top             =   2295
         Width           =   5595
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   45
         TabIndex        =   7
         Top             =   1125
         Width           =   5595
      End
      Begin VB.ListBox lstLocation 
         Height          =   450
         ItemData        =   "bootfiles.frx":038A
         Left            =   45
         List            =   "bootfiles.frx":03A9
         TabIndex        =   6
         Top             =   1395
         Width           =   5595
      End
      Begin VB.ListBox lstRegister 
         Height          =   450
         ItemData        =   "bootfiles.frx":043A
         Left            =   45
         List            =   "bootfiles.frx":0450
         TabIndex        =   5
         Top             =   1845
         Width           =   5595
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   45
         TabIndex        =   4
         Top             =   2565
         Width           =   5595
      End
      Begin VB.TextBox txtSize 
         Height          =   285
         Left            =   45
         TabIndex        =   3
         Top             =   2835
         Width           =   5595
      End
      Begin VB.TextBox txtVersion 
         Height          =   285
         Left            =   45
         TabIndex        =   2
         Top             =   3105
         Width           =   5595
      End
      Begin MSComctlLib.ListView lvwFiles 
         Height          =   960
         Left            =   45
         TabIndex        =   0
         Top             =   180
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   1693
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   600
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "File"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Location"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Register"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Shared"
            Object.Width           =   485
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Size"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Version"
            Object.Width           =   1940
         EndProperty
      End
   End
End
Attribute VB_Name = "Files"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sLength(30) As String

Private Sub chkShared_Click()
   If chkShared.Value = 1 Then
      lvwFiles.SelectedItem.ListSubItems(4).Text = "$(Shared)"
   Else
      lvwFiles.SelectedItem.ListSubItems(4).Text = ""
   End If
End Sub

Private Function CombineLine(Index)
Dim sTemp As String
Dim X As Integer
   sTemp = lvwFiles.ListItems(Index).ListSubItems(1).Text
   For X = 1 To 7
      sTemp = sTemp & "," & lvwFiles.ListItems(Index).ListSubItems(X).Text
   Next
   CombineLine = sTemp
End Function

Private Sub cmdApply_Click()
Dim I As Integer

   For I = 1 To lvwFiles.ListItems.Count - 1
      UpdateClean "File" & I, sLength(I), CombineLine(I)
   Next I
   
   Unload Me
End Sub

Private Sub Form_Load()

   Dim sLine As String
   Dim sTmp As String
   Dim iTmp As Integer
   Dim ID As Integer
   Dim iFile As Integer

   iFile = FreeFile
   
   Open sFile For Input As iFile
         
         Do Until sLine = "[" & sFiles & "]"
            Line Input #iFile, sLine
         Loop
         ID = 0
         Do
            ID = ID + 1
            
            Line Input #iFile, sLine
            
            If sLine <> "" Then
            
               sLength(ID) = Left(sLine, Len(sLine) - 6)
            
               sLine = Right(sLine, Len(sLine) - 4)
               iTmp = InStr(1, sLine, "=")
               sTmp = Left(sLine, iTmp - 1)
               lvwFiles.ListItems.Add , , sTmp
               sLine = Right(sLine, Len(sLine) - iTmp)
               
               iTmp = InStr(1, sLine, ",")
               sTmp = Left(sLine, iTmp - 1)
               lvwFiles.ListItems(ID).ListSubItems.Add , , sTmp
               sLine = Right(sLine, Len(sLine) - iTmp)
               
               iTmp = InStr(1, sLine, ",")
               sTmp = Left(sLine, iTmp - 1)
               lvwFiles.ListItems(ID).ListSubItems.Add , , sTmp
               sLine = Right(sLine, Len(sLine) - iTmp)
               
               iTmp = InStr(1, sLine, ",")
               sTmp = Left(sLine, iTmp - 1)
               lvwFiles.ListItems(ID).ListSubItems.Add , , sTmp
               sLine = Right(sLine, Len(sLine) - iTmp)
               
               iTmp = InStr(1, sLine, ",")
               sTmp = Left(sLine, iTmp - 1)
               lvwFiles.ListItems(ID).ListSubItems.Add , , sTmp
               sLine = Right(sLine, Len(sLine) - iTmp)
               
               iTmp = InStr(1, sLine, ",")
               sTmp = Left(sLine, iTmp - 1)
               lvwFiles.ListItems(ID).ListSubItems.Add , , sTmp
               sLine = Right(sLine, Len(sLine) - iTmp)
               
               iTmp = InStr(1, sLine, ",")
               sTmp = Left(sLine, iTmp - 1)
               lvwFiles.ListItems(ID).ListSubItems.Add , , sTmp
               sLine = Right(sLine, Len(sLine) - iTmp)
               
               lvwFiles.ListItems(ID).ListSubItems.Add , , sLine
            
            End If
         
         Loop Until sLine = ""
         
   Close iFile
End Sub

Private Sub Form_Resize()
On Error Resume Next

   frmBoot.Width = Files.ScaleWidth
   frmBoot.Height = Files.ScaleHeight - 300
   
   cmdHelp.Left = frmBoot.Width - 200
   cmdHelp.Top = frmBoot.Height + 50
   
   cmdApply.Left = Files.ScaleWidth - 200 - 1260
   cmdApply.Top = frmBoot.Height + 50
   
   lvwFiles.Width = frmBoot.Width - 100
   lvwFiles.Height = frmBoot.Height - 2550
   
   txtFile.Width = frmBoot.Width - 100
   lstLocation.Width = frmBoot.Width - 100
   lstRegister.Width = frmBoot.Width - 100
   chkShared.Width = frmBoot.Width - 100
   txtDate.Width = frmBoot.Width - 100
   txtSize.Width = frmBoot.Width - 100
   txtVersion.Width = frmBoot.Width - 100
   
   txtFile.Top = lvwFiles.Height + 150
   lstLocation.Top = txtFile.Top + txtFile.Height
   lstRegister.Top = lstLocation.Top + lstLocation.Height
   chkShared.Top = lstRegister.Top + lstRegister.Height
   txtDate.Top = chkShared.Top + chkShared.Height
   txtSize.Top = txtDate.Top + txtDate.Height
   txtVersion.Top = txtSize.Top + txtSize.Height

End Sub

Private Sub lstLocation_DblClick()
   lstLocation_LostFocus
End Sub

Private Sub lstLocation_LostFocus()
   lvwFiles.SelectedItem.ListSubItems(2).Text = lstLocation.List(lstLocation.ListIndex)
End Sub

Private Sub lstRegister_DblClick()
   lstRegister_LostFocus
End Sub

Private Sub lstRegister_LostFocus()
   lvwFiles.SelectedItem.ListSubItems(3).Text = lstRegister.List(lstRegister.ListIndex)
End Sub

Private Sub lvwFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)

   txtVersion.Text = Item.ListSubItems(7).Text
   txtSize.Text = Item.ListSubItems(6).Text & " bytes"
   txtDate.Text = Item.ListSubItems(5).Text
   txtFile.Text = Item.ListSubItems(1).Text
   
   If Left(Item.ListSubItems(2).Text, 1) = "$" Then
      lstLocation.Text = Item.ListSubItems(2).Text
   Else
      lstLocation.Text = "[Path]"
   End If
   If Left(Item.ListSubItems(3).Text, 1) = "$" Then
      lstRegister.Text = Item.ListSubItems(3).Text
   ElseIf Not Left(Item.ListSubItems(3).Text, 1) = "" Then
      lstRegister.Text = "[FileName]"
   End If
   
   If Item.ListSubItems(4).Text = "" Then
      chkShared.Value = 0
   Else
      chkShared.Value = 1
   End If
   
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtDate_LostFocus
   End If
End Sub

Private Sub txtDate_LostFocus()
   lvwFiles.SelectedItem.ListSubItems(5).Text = txtDate.Text
End Sub

Private Sub txtFile_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtFile_LostFocus
   End If
End Sub

Private Sub txtFile_LostFocus()
   lvwFiles.SelectedItem.ListSubItems(1).Text = txtFile.Text
End Sub

Private Sub txtSize_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtSize_LostFocus
   End If
End Sub

Private Sub txtSize_LostFocus()
   lvwFiles.SelectedItem.ListSubItems(1).Text = txtSize.Text
End Sub

Private Sub txtVersion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      txtVersion_LostFocus
   End If
End Sub

Private Sub txtVersion_LostFocus()
   lvwFiles.SelectedItem.ListSubItems(1).Text = txtVersion.Text
End Sub
