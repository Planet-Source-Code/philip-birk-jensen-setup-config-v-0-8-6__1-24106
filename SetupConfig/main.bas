Attribute VB_Name = "modMain"
Option Explicit
Option Compare Binary

Private Type BrowseInfo
   hWndOwner As Long
   pidlRoot As Long
   sDisplayName As String
   sTitle As String
   ulFlags As Long
   lpfn As Long
   lParam As Long
   iImage As Long
End Type

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (bBrowse As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal lItem As Long, ByVal sDir As String) As Long

Global sFile As String
Global sVBProject As String
Global sVBPath As String
Global sVBEXE As String
Global sFiles As String


Public Sub UpdateTitle(Unit As String, OldText As String, Text As String)
Dim iTmp
   With Clean
      iTmp = InStr(1, .txtData, "[Setup]")
      iTmp = InStr(iTmp, .txtData, Unit)
      .txtData.SelStart = iTmp + Len(Unit)
      .txtData.SelLength = Len(OldText)
      .txtData.SelText = Text
   End With

End Sub

Public Sub UpdateClean(Unit As String, OldText As String, Text As String)
Dim iTmp
   With Clean
      iTmp = InStr(1, .txtData, Unit)
      .txtData.SelStart = iTmp + Len(Unit)
      .txtData.SelLength = Len(OldText)
      .txtData.SelText = Text
   End With

End Sub

Public Function BDirectory(frm As Form) As String
Dim bri As BrowseInfo
Dim Item As Long
Dim Path As String
   
   bri.hWndOwner = frm.hwnd
   bri.pidlRoot = 0
   bri.sDisplayName = Space$(260)
   bri.sTitle = "Select Temp Directory"
   bri.ulFlags = 1 ' Return directory name.
   bri.lpfn = 0
   bri.lParam = 0
   bri.iImage = 0
   
   Item = SHBrowseForFolder(bri)
   If Item Then
       Path = Space$(260)
       If SHGetPathFromIDList(Item, Path) Then
           BDirectory = Left(Path, InStr(Path, Chr$(0)) - 1)
       Else
           BDirectory = ""
       End If
   End If
End Function


Public Sub FileChanges()

   With mdiMain
      .sbrInfo.Panels(2).Text = Round(FileLen(sFile) / 1024, 3) & " kb"
      .sbrInfo.Panels(3).Text = FileDateTime(sFile)
   End With

End Sub

Public Sub FileLoaded(File As Integer)
   
   Select Case File
      Case 0 ' setup.lst
         With mdiMain
            .mProject(0).Enabled = True
            .mProject(1).Enabled = True
            .mProject(2).Enabled = True
            .mProject(4).Enabled = True
            .mProject(6).Enabled = True
            .tbrMenu.Buttons(5).Enabled = True
            .tbrMenu.Buttons(6).Enabled = True
            .tbrMenu.Buttons(7).Enabled = True
            .tbrMenu.Buttons(8).Enabled = True
            .sbrInfo.Panels(1).Text = sFile
         End With
         FileChanges
      Case 1 ' Visual Basic project
         With mdiMain
            .mVB(2).Enabled = True
            .mVB(3).Enabled = True
            .mVB(4).Enabled = True
            .mVB(6).Enabled = True
            '.mVB(8).Enabled = True
            .tbrMenu.Buttons(10).ButtonMenus(3).Enabled = True
            .tbrMenu.Buttons(10).ButtonMenus(4).Enabled = True
            .tbrMenu.Buttons(10).ButtonMenus(5).Enabled = True
            .tbrMenu.Buttons(10).ButtonMenus(7).Enabled = True
            '.tbrMenu.Buttons(10).ButtonMenus(9).Enabled = True
         End With
   End Select
   
End Sub

Public Sub VBOpen()

   'ShellExecute mdiMain.hwnd, "open", sVBEXE, " " & sVBProject, "", 1
   Shell sVBEXE & " """ & sVBProject & """", vbNormalFocus
   
End Sub

Public Sub VBCompile()
   
   'ShellExecute mdiMain.hwnd, "open", sVBEXE, sVBProject & " /m", "", 1
   Shell sVBEXE & " """ & sVBProject & """ /make", vbNormalFocus
   
End Sub

Public Sub VBRunP()

   'ShellExecute mdiMain.hwnd, "open", sVBEXE, sVBProject & " /r", "", 1
   Shell sVBEXE & " """ & sVBProject & """ /runexit", vbNormalFocus
   
End Sub

Public Sub VBDir()
Dim iTmp As Integer
Dim sTmp As String

   iTmp = InStrRev(sVBProject, "\")
   
   sTmp = Left(sVBProject, iTmp)

   ShellExecute mdiMain.hwnd, "explore", sTmp, "", "", 1
   
End Sub


Public Function CutRight(Text As String)
Dim iTmp As Integer

   iTmp = InStr(1, Text, "=")
   iTmp = Len(Text) - iTmp
   CutRight = Right(Text, iTmp)
   
End Function

Public Function CutLeft(Text As String)
On Error Resume Next
Dim iTmp As Integer

   iTmp = InStr(1, Text, "=") - 1
   CutLeft = Left(Text, iTmp)
   
End Function
