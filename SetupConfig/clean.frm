VERSION 5.00
Begin VB.Form Clean 
   Caption         =   "Form1"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3405
   Icon            =   "clean.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2805
   ScaleWidth      =   3405
   Begin VB.TextBox txtData 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "clean.frx":038A
      Top             =   0
      Width           =   2310
   End
End
Attribute VB_Name = "Clean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Op As Boolean

Private Sub Form_Load()
Dim iFile As Integer

   iFile = FreeFile
   
   Op = True

   Open sFile For Input As iFile
      Screen.MousePointer = 11
      txtData.Text = Input$(LOF(iFile), iFile)
      Screen.MousePointer = 0
   Close iFile
   
   Clean.Caption = "Clean"
   
   Op = True
   
End Sub

Private Sub Form_Resize()

   txtData.Height = Clean.ScaleHeight
   txtData.Width = Clean.ScaleWidth
   
End Sub

Private Sub txtData_Change()
Dim iFile As Integer
   If Op = False Then
      iFile = FreeFile
      
      Open sFile For Binary As iFile
         'Print iFile, txtData.Text
         Put iFile, , txtData.Text
      Close iFile
   Else
      Op = False
   End If
   
   
   
End Sub
