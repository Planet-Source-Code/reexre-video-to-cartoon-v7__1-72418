Attribute VB_Name = "modINI"
' ModINIWG.bas

Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
 (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
 ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
 (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
 ByVal lpDefault As String, ByVal lpReturnedString As String, _
 ByVal nSize As Long, ByVal lpFileName As String) As Long
 ' lpDefault is the string return if no ini file found.
'------------------------------------------------------

' File  stuff  ini & RecentFiles
Public IniTitle$, IniSpec$

'Public Const MaxRecentFiles = 9 'Max of files to show in list
'Public NumRecentFiles As Long
'Public RecentFiles$()

Public Function WriteINI(Title$, TheKey$, Info$, FileSpec$) As Boolean
   WritePrivateProfileString Title$, TheKey$, Info$, FileSpec$
End Function

Public Function GetINI(Title$, TheKey$, ret$, FileSpec$) As String
Dim n As Long
   On Error GoTo NoINI
   ret$ = String(255, 0)
   n = GetPrivateProfileString(Title$, TheKey$, "", ret$, 255, FileSpec$)
   'N is the number of characters copied to Ret$
   If n <> 0 Then
'     GetINI = True
     ret$ = Left$(ret$, n)
   Else
'     GetINI = False
     ret$ = ""
   End If
   GetINI = ret$

   On Error GoTo 0
   Exit Function
'==========
NoINI:
'GetINI = False
ret$ = ""
GetINI = ret$

End Function



