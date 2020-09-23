VERSION 5.00
Begin VB.Form Tip 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tips"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picFrame 
      BackColor       =   &H00FFFFFF&
      Height          =   3075
      Left            =   0
      Picture         =   "tip.frx":0000
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   0
      Top             =   0
      Width           =   3075
      Begin VB.Label lblHead 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tips"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   450
         TabIndex        =   5
         Top             =   45
         Width           =   2535
      End
      Begin VB.Line Line4 
         X1              =   16
         X2              =   120
         Y1              =   164
         Y2              =   164
      End
      Begin VB.Line Line3 
         X1              =   16
         X2              =   168
         Y1              =   143
         Y2              =   143
      End
      Begin VB.Label lblDone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Done "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2520
         MousePointer    =   2  'Cross
         TabIndex        =   4
         Top             =   2025
         Width           =   510
      End
      Begin VB.Label lblStop 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stop Showing "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1800
         MousePointer    =   2  'Cross
         TabIndex        =   3
         Top             =   2340
         Width           =   1230
      End
      Begin VB.Label lblNext 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Next Tip "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2250
         MousePointer    =   2  'Cross
         TabIndex        =   2
         Top             =   2655
         Width           =   780
      End
      Begin VB.Line Line2 
         X1              =   16
         X2              =   150
         Y1              =   185
         Y2              =   185
      End
      Begin VB.Line Line1 
         X1              =   16
         X2              =   16
         Y1              =   201
         Y2              =   31
      End
      Begin VB.Label lblTip 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"tip.frx":048E
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   360
         TabIndex        =   1
         Top             =   420
         Width           =   2580
      End
   End
End
Attribute VB_Name = "Tip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Info As New Collection
Dim sTemp As String, sLast As String

Private Sub Form_Load()
Randomize

   picFrame.BackColor = RGB(240, 240, 240)
   lblDone.BackColor = RGB(255, 255, 230)
   lblStop.BackColor = RGB(255, 255, 230)
   lblNext.BackColor = RGB(255, 255, 230)
   
   Open App.Path & "\tips.txt" For Input As #1
      Do Until EOF(1)
         Line Input #1, sTemp
         Info.Add sTemp
      Loop
   Close #1
   
   lblTip.Caption = Info.Item(Int(Rnd * Info.Count) + 1)
   sLast = Info.Item(Int(Rnd * Info.Count) + 1)

End Sub

Private Sub lblDone_Click()

   Unload Me

End Sub

Private Sub lblNext_Click()
Randomize
   Do
      lblTip.Caption = Info.Item(Int(Rnd * Info.Count) + 1)
   Loop Until sLast <> lblTip.Caption
   sLast = lblTip.Caption
End Sub

Private Sub lblStop_Click()
   SaveSetting "SetupConfig", "StartUp", "ShowTip", "0"
   Unload Me
End Sub
