VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'Public IniTitle$, IniSpec$

Public FileName$


'Public Const MaxRecentFiles = 9 'Max of files to show in list
'Public NumRecentFiles As Long
'Public RecentFiles$()

Public Function WriteINI(Title$, TheKey$, Value$) As Boolean
WritePrivateProfileString Title$, TheKey$, Value$, FileName$
End Function

Public Function GetINI(Title$, TheKey$) As String
Dim n As Long
Dim Ret$
On Error GoTo NoINI
Ret$ = String(255, 0)
n = GetPrivateProfileString(Title$, TheKey$, "", Ret$, 255, FileName$)
'N is the number of characters copied to Ret$
If n <> 0 Then
    '     GetINI = True
    Ret$ = Left$(Ret$, n)
Else
    '     GetINI = False
    Ret$ = "0"
End If
GetINI = Ret$

On Error GoTo 0
Exit Function
'==========
NoINI:
'GetINI = False
Ret$ = ""
GetINI = Ret$

End Function


