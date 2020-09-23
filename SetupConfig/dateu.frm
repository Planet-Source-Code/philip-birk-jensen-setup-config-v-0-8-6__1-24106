VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form DateU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Date Utility"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "dateu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4725
   Begin VB.CommandButton cmdFind 
      Caption         =   "Auto Add"
      Height          =   285
      Left            =   900
      TabIndex        =   10
      Top             =   1575
      Width           =   870
   End
   Begin MSComDlg.CommonDialog cdlAdd 
      Left            =   3330
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Add File"
      Filter          =   "Applications|*.exe|Setup.lst|setup.lst|Text Files|*.txt|Cab Files|*.cab|All Files|*.*"
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add File"
      Height          =   285
      Left            =   45
      TabIndex        =   9
      Top             =   1575
      Width           =   870
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
      Left            =   4500
      TabIndex        =   8
      Top             =   1665
      Width           =   195
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   285
      Left            =   3555
      TabIndex        =   7
      Top             =   1575
      Width           =   780
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   285
      Left            =   2745
      TabIndex        =   6
      Top             =   1575
      Width           =   735
   End
   Begin VB.Frame frmFiles 
      Caption         =   "Files"
      Height          =   1455
      Left            =   2655
      TabIndex        =   4
      Top             =   45
      Width           =   2040
      Begin VB.ListBox lstFiles 
         Height          =   1050
         IntegralHeight  =   0   'False
         Left            =   135
         OLEDropMode     =   1  'Manual
         TabIndex        =   5
         Top             =   270
         Width           =   1770
      End
   End
   Begin VB.Frame frmModified 
      Caption         =   "Modified"
      Height          =   690
      Left            =   0
      TabIndex        =   2
      Top             =   810
      Width           =   2580
      Begin VB.TextBox txtModifiedT 
         Height          =   285
         Left            =   1350
         TabIndex        =   12
         Top             =   270
         Width           =   1095
      End
      Begin VB.TextBox txtModifiedD 
         Height          =   285
         Left            =   135
         TabIndex        =   3
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Frame frmCreated 
      Caption         =   "Created"
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   2580
      Begin VB.TextBox txtCreatedT 
         Height          =   285
         Left            =   1350
         TabIndex        =   11
         Top             =   270
         Width           =   1095
      End
      Begin VB.TextBox txtCreatedD 
         Height          =   285
         Left            =   135
         TabIndex        =   1
         Top             =   270
         Width           =   1095
      End
   End
End
Attribute VB_Name = "DateU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Most of this code, is taken from my other project "laTEXT"
' So the code might be messy, and have some not needed code in it,
' I will try fixing this for next version(and then next version agin, and so on)
'
' The auto detection will detect more in the next version

Option Explicit

Const OF_READWRITE = &H2
Const OF_READ = &H0

Const OFS_MAXPATHNAME = 128

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

Private Type File
   sFileName   As String ' path
   tCreate     As String ' time
   tModify     As String
   tAccess     As String
   dCreate     As String ' date
   dModify     As String
   dAccess     As String
   lAttr       As Long   ' Attributes
End Type

Dim fCurrent As File
Dim ReadOnly As Byte
Dim bDTSelect As Integer
Dim test As Long
Dim cFiles As New Collection


Private Sub SetFileInfo()
Dim oStructure As OFSTRUCT
Dim fTime(2) As FILETIME
Dim sTime(2) As SYSTEMTIME
Dim hFile As Long
Dim I As Byte

   With fCurrent
      
      hFile = OpenFile(.sFileName, oStructure, OF_READWRITE)
      
      sTime(0) = ConvertTimeDate(.tCreate, .dCreate)
      sTime(1) = ConvertTimeDate(.tAccess, .dAccess)
      sTime(2) = ConvertTimeDate(.tModify, .dModify)
      
      For I = 0 To 2
         SystemTimeToFileTime sTime(I), fTime(I)
         LocalFileTimeToFileTime fTime(I), fTime(I)
      Next
      
      SetFileTime hFile, fTime(0), fTime(1), fTime(2)
      
   End With
   
   CloseHandle hFile
End Sub

Private Function ConvertTimeDate(sTime As String, sDate As String) As SYSTEMTIME
   ConvertTimeDate.wDay = Mid(sDate, 1, 2)
   ConvertTimeDate.wMonth = Mid(sDate, 4, 2)
   ConvertTimeDate.wYear = Mid(sDate, 7, 4)
   
   ConvertTimeDate.wHour = Mid(sTime, 1, 2)
   ConvertTimeDate.wMinute = Mid(sTime, 4, 2)
   ConvertTimeDate.wSecond = Mid(sTime, 7, 2)
   ConvertTimeDate.wMilliseconds = Mid(sTime, 10, 3)
End Function

Private Sub GetFileInfo()

Dim oStructure As OFSTRUCT
Dim fTime(2) As FILETIME
Dim sTime(2) As SYSTEMTIME
Dim hFile As Long
Dim I As Byte

   
   With fCurrent
      
      .lAttr = GetAttr(.sFileName)

      hFile = OpenFile(.sFileName, oStructure, OF_READ)
      
      GetFileTime hFile, fTime(0), fTime(1), fTime(2)
      
      For I = 0 To 2
         FileTimeToLocalFileTime fTime(I), fTime(I)
         FileTimeToSystemTime fTime(I), sTime(I)
      Next
      
      .dCreate = Format(sTime(0).wDay, "00") & "-" & Format(sTime(0).wMonth, "00") & "-" & sTime(0).wYear
      .dAccess = Format(sTime(1).wDay, "00") & "-" & Format(sTime(1).wMonth, "00") & "-" & sTime(1).wYear
      .dModify = Format(sTime(2).wDay, "00") & "-" & Format(sTime(2).wMonth, "00") & "-" & sTime(2).wYear
      
      .tCreate = Format(sTime(0).wHour, "00") & ":" & Format(sTime(0).wMinute, "00") & ":" & Format(sTime(0).wSecond) & "," & Format(sTime(0).wMilliseconds, "000")
      .tAccess = Format(sTime(1).wHour, "00") & ":" & Format(sTime(1).wMinute, "00") & ":" & Format(sTime(1).wSecond) & "," & Format(sTime(1).wMilliseconds, "000")
      .tModify = Format(sTime(2).wHour, "00") & ":" & Format(sTime(2).wMinute, "00") & ":" & Format(sTime(2).wSecond) & "," & Format(sTime(2).wMilliseconds, "000")
      
   
      txtCreatedD.Text = .dCreate
      txtModifiedD.Text = .dModify
      
      txtCreatedT.Text = .tCreate
      txtModifiedT.Text = .tModify
      
   End With
   CloseHandle hFile
End Sub


Private Sub cmdAdd_Click()
On Error GoTo CancelError:
   cdlAdd.ShowOpen
   lstFiles.AddItem Right(cdlAdd.FileName, Len(cdlAdd.FileName) - InStrRev(cdlAdd.FileName, "\"))
   cFiles.Add cdlAdd.FileName, Right(cdlAdd.FileName, Len(cdlAdd.FileName) - InStrRev(cdlAdd.FileName, "\"))
CancelError:
End Sub

Private Sub cmdApply_Click()
   With fCurrent
      .dCreate = txtCreatedD.Text
      .dModify = txtModifiedD.Text
      .tCreate = txtCreatedT.Text
      .tModify = txtModifiedT.Text
   End With
   SetFileInfo
End Sub

Private Sub cmdDone_Click()
   Unload Me
End Sub

Private Sub cmdFind_Click()
Dim sTemp As String

   lstFiles.AddItem Right(sFile, Len(sFile) - InStrRev(sFile, "\"))
   cFiles.Add sFile, Right(sFile, Len(sFile) - InStrRev(sFile, "\"))
   
   sTemp = Left(sFile, InStrRev(sFile, "\")) & "\setup.exe"
   
   lstFiles.AddItem Right(sTemp, Len(sTemp) - InStrRev(sTemp, "\"))
   cFiles.Add sTemp, Right(sTemp, Len(sTemp) - InStrRev(sTemp, "\"))
   
End Sub

Private Sub lstFiles_Click()
   fCurrent.sFileName = cFiles.Item(lstFiles.List(lstFiles.ListIndex))
   GetFileInfo
End Sub

Private Sub lstFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo NoFiles:
Dim sTemp

   For Each sTemp In Data.Files
      lstFiles.AddItem Right(sTemp, Len(sTemp) - InStrRev(sTemp, "\"))
      cFiles.Add sTemp, Right(sTemp, Len(sTemp) - InStrRev(sTemp, "\"))
   Next
   
NoFiles:
End Sub
