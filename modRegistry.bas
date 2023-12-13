Attribute VB_Name = "modRegistry"
Option Explicit

Public Declare Function RegCreateKeyEx Lib "Coredll" Alias "RegCreateKeyExW" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegCloseKey Lib "Coredll" (ByVal hKey As Long) As Long
Public Declare Function RegSetValueEx Lib "Coredll" Alias "RegSetValueExW" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const HKEY_CLASSES_ROOT = &H80000000

Public Function CreateNewKey(sNewKeyName As String)
    Dim hKeyNew As Long
    Dim lRetVal As Long
    lRetVal = RegCreateKeyEx(HKEY_CLASSES_ROOT, sNewKeyName, CLng(0), vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, CLng(0), hKeyNew, lRetVal)
    RegCloseKey (hKeyNew)
End Function

Public Function SaveKeyValue(sNewKeyName As String, sValue As String)
    Dim lRetVal As Long
    lRetVal = RegSetValueEx(HKEY_CLASSES_ROOT, sNewKeyName, CLng(0), CLng(0), sValue, CLng(0))
    RegCloseKey (
End Function
