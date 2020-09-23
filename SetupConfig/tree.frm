VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Tree 
   Caption         =   "Explore"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   Icon            =   "tree.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   8925
   Begin MSComctlLib.ImageList lstSmall 
      Left            =   5895
      Top             =   4545
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tree.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tree.frx":0724
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList lstLarge 
      Left            =   4455
      Top             =   4275
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tree.frx":0ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tree.frx":1398
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLocation 
      Height          =   315
      Left            =   3285
      TabIndex        =   2
      Top             =   0
      Width           =   960
   End
   Begin MSComctlLib.ListView lvwDetails 
      Height          =   2670
      Left            =   3285
      TabIndex        =   1
      Top             =   315
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   4710
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "lstLarge"
      SmallIcons      =   "lstSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Option"
         Object.Width           =   3087
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Data"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwSetup 
      Height          =   5685
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   10028
      _Version        =   393217
      Indentation     =   441
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
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
End
Attribute VB_Name = "Tree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iSelect As Integer
Dim sLocation As String
Dim dblClick As Boolean
Dim nClick As Boolean

Private Sub cmdCancel_Click()
   
   Unload Me
   
End Sub

Private Sub Form_Load()
Dim sLine As String
Dim sTmp As String
Dim ID As Integer
Dim iFile As Integer

   iFile = FreeFile
   
   If Len(sFile) > 28 Then
      sTmp = Left(sFile, 25) & "...\setup.lst"
   Else
      sTmp = sFile
   End If
   
   tvwSetup.Nodes.Add , , , sTmp

   Open sFile For Input As iFile
      Do Until EOF(iFile)
      
         Line Input #iFile, sLine
         
         sTmp = Left(sLine, 1)
         
         Select Case sTmp
            Case "["
               tvwSetup.Nodes.Add 1, tvwChild, , sLine
               ID = tvwSetup.Nodes.Count
               tvwSetup.Nodes(ID).Bold = True
            Case ";"
               
            Case ""
            Case Else
               sTmp = CutLeft(sLine)
               tvwSetup.Nodes.Add ID, tvwChild, , sTmp
         End Select
         
      Loop
   Close iFile
   Location (tvwSetup.Nodes(2).Text)
   tvwSetup.Nodes(2).Selected = True
dblClick = False

End Sub

Private Sub Form_Resize()
On Error Resume Next
   tvwSetup.Height = Tree.ScaleHeight
   lvwDetails.Height = Tree.ScaleHeight - 300
   lvwDetails.Width = Tree.ScaleWidth - 3275
   txtLocation.Width = Tree.ScaleWidth - 3275
End Sub

Private Sub lvwDetails_Click()
   
   nClick = False
   
End Sub

Private Sub lvwDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)

   iSelect = Item.Index
   txtLocation.Text = Item.ListSubItems(1).Text
   nClick = True

End Sub

Private Sub lvwDetails_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = 2 Then
      If nClick = True Then
         mdiMain.mList(2).Enabled = True
      Else
         mdiMain.mList(2).Enabled = False
      End If
      Tree.PopupMenu mdiMain.mnuList
   End If

End Sub

Private Sub tvwSetup_Click()

   dblClick = False
   
End Sub

Private Sub tvwSetup_DblClick()
   
   dblClick = True
   
End Sub

Private Sub tvwSetup_KeyDown(KeyCode As Integer, Shift As Integer)
   
   dblClick = False

End Sub

Private Sub tvwSetup_NodeClick(ByVal Node As MSComctlLib.Node)
Dim nodeFind

   If "[" = Left(Node.Text, 1) Then
      lvwDetails.ListItems.Clear
      Location Node.Text
   ElseIf Node.Index <> 1 Then
      If sLocation <> Node.Parent Then
         lvwDetails.ListItems.Clear
         Location Node.Parent
      End If
      
         Set nodeFind = lvwDetails.FindItem(Node.Text, 0, , 0)
         nodeFind.EnsureVisible
         nodeFind.Selected = True
         iSelect = nodeFind.Index
         txtLocation.Text = nodeFind.ListSubItems(1).Text
         
      If dblClick = True Then
         lvwDetails.SetFocus
      End If
   End If
   
End Sub

Private Sub Location(Text As String)
Dim sLine As String
Dim sTmp As String
Dim ID As Integer
Dim iFile As Integer

   iFile = FreeFile

   Open sFile For Input As iFile
      Do Until EOF(iFile)
      
         Line Input #iFile, sLine
         
         If sLine = Text Then
            
            Do Until EOF(iFile)
               Line Input #iFile, sLine
               If sLine = "" Then
                  Exit Do
               End If
               sTmp = CutLeft(sLine)
               lvwDetails.ListItems.Add , , sTmp, 2, 2
               sTmp = CutRight(sLine)
               lvwDetails.ListItems(lvwDetails.ListItems.Count) _
               .ListSubItems.Add , , sTmp
            Loop
            GoTo ExitLoop
         End If
         
      Loop

ExitLoop:
   sLocation = Text
   Close iFile
End Sub

'Private Sub Location(Text As String)
'Dim sLine As String
'Dim sTmp As String
'Dim ID As Integer
'Dim iFile As Integer
'
'   iFile = FreeFile
'
'   Open sFile For Input As iFile
'      Do Until EOF(iFile)
'
'         Line Input #iFile, sLine
'
'         If sLine = Text Then
'
'            Do Until EOF(iFile)
'               Line Input #iFile, sLine
'               If sLine = "" Then
'                  Exit Do
'               End If
'               sTmp = CutLeft(sLine)
'               lvwDetails.ListItems.Add , , sTmp, 2, 2
'               sTmp = CutRight(sLine)
'               lvwDetails.ListItems(lvwDetails.ListItems.Count) _
'               .ListSubItems.Add , , sTmp
'            Loop
'            GoTo ExitLoop
'         End If
'
'      Loop
'
'ExitLoop:
'   sLocation = Text
'   Close iFile
'End Sub

Private Sub txtLocation_KeyUp(KeyCode As Integer, Shift As Integer)
Dim iTmp As Integer
   If KeyCode = 13 Then
      UpdateClean lvwDetails.SelectedItem.Text, lvwDetails.SelectedItem.SubItems(1), txtLocation.Text
'      With Clean
'         iTmp = InStr(1, .txtData.Text, lvwDetails.SelectedItem.Text)
'         .txtData.SelStart = iTmp + Len(lvwDetails.SelectedItem.Text)
'         .txtData.SelLength = Len(lvwDetails.SelectedItem.SubItems(1))
'         .txtData.SelText = txtLocation.Text
'      End With
      
      lvwDetails.ListItems(iSelect).ListSubItems(1).Text = txtLocation.Text
   End If
End Sub
