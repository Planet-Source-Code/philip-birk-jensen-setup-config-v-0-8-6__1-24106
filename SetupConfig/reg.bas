Attribute VB_Name = "modReg"
' --------------------------------------------------------------------
'  I haven't written the following code, and can't remember who did
' --------------------------------------------------------------------

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const SW_SHOWNORMAL = 1
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Const ERROR_SUCCESS = 0&
Const REG_SZ = 1
Const REG_DWORD = 4

Sub SaveKey(Hkey As Long, strPath As String)
   On Error Resume Next
   Dim KeyHand As Long
   Call RegCreateKey(Hkey, strPath, KeyHand&)
   Call RegCloseKey(KeyHand&)
End Sub

Function GetString(Hkey As Long, strPath As String, strValue As String)
   On Error Resume Next
   Dim KeyHand As Long
   Dim datatype As Long
   Dim lResult As Long
   Dim strBuf As String
   Dim lDataBufSize As Long
   Dim lValueType As Long
   Dim intZeroPos As Integer
   Call RegOpenKey(Hkey, strPath, KeyHand)
   lResult = RegQueryValueEx(KeyHand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
   If lValueType = REG_SZ Then
      strBuf = String(lDataBufSize, " ")
      lResult = RegQueryValueEx(KeyHand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
      If lResult = ERROR_SUCCESS Then
         intZeroPos = InStr(strBuf, Chr$(0))
         If intZeroPos > 0 Then
            GetString = Left$(strBuf, intZeroPos - 1)
         Else
            GetString = strBuf
         End If
      End If
   End If
End Function

Sub SaveString(Hkey As Long, strPath As String, strValue As String, strData As String)
   On Error Resume Next
   Dim KeyHand As Long
   Call RegCreateKey(Hkey, strPath, KeyHand)
   Call RegSetValueEx(KeyHand, strValue, 0, REG_SZ, ByVal strData, Len(strData))
   Call RegCloseKey(KeyHand)
End Sub

Function GetDword(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
   On Error Resume Next
   Dim lResult As Long
   Dim lValueType As Long
   Dim lBuf As Long
   Dim lDataBufSize As Long
   Dim KeyHand As Long
   Call RegOpenKey(Hkey, strPath, KeyHand)
   lDataBufSize = 4
   lResult = RegQueryValueEx(KeyHand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
   If lResult = ERROR_SUCCESS Then
      If lValueType = REG_DWORD Then
         GetDword = lBuf
      End If
   End If
   Call RegCloseKey(KeyHand)
End Function

Sub SaveDword(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
   On Error Resume Next
   Dim lResult As Long
   Dim KeyHand As Long
   Call RegCreateKey(Hkey, strPath, KeyHand)
   lResult = RegSetValueEx(KeyHand, strValueName, 0&, REG_DWORD, lData, 4)
   Call RegCloseKey(KeyHand)
End Sub

Sub DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
   On Error Resume Next
   Dim KeyHand As Long
   Call RegOpenKey(Hkey, strPath, KeyHand)
   Call RegDeleteValue(KeyHand, strValue)
   Call RegCloseKey(KeyHand)
End Sub

Sub DeleteKey(ByVal Hkey As Long, ByVal strKey As String)
   Dim r As Long
   Call RegDeleteKey(Hkey, strKey)
End Sub


