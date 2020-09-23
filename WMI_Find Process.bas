Attribute VB_Name = "WMI_FindProcess"
'***************************************************************************
'
' Function checks if a process is running on your computer
' Based on the WMI (Windows Management Instrurmentation) code on MSDN
'
' Mark Mokoski
' 15-APR-2005
' www.rjillc.com
'
' This Function requires WMI (Windows Management Instrurmentation).
' WMI is part of Windows 2000, XP
'
' For more information see the MSDN Web site
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/wmi_tasks__processes.asp
'
' ****************************************************************************

Option Explicit
Public ProcessArray()            As String
Public ProcArraySize             As Integer

Public Function IsProcessRunning(strProcess As String)

' Function checks if a process is running on your computer
' Based on the WMI (Windows Management Instrurmentation) code on MSDN
' This Function requires WMI (Windows Management Instrurmentation).
' WMI is part of Windows 2000, XP
' For more information see the MSDN Web site
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/wmi_tasks__processes.asp
'
'   Function returns Boolean value (True if task running)

Dim ProcessEXE            As Object

IsProcessRunning = False

strProcess = UCase(strProcess)
For Each ProcessEXE In GetObject("winmgmts://").InstancesOf("win32_process")
    
    'Force uppercase for name compair to avoid typed case problems (ex. Camel typed names)
    If UCase(ProcessEXE.Name) = strProcess Then
        IsProcessRunning = True
        Exit Function
    End If
Next


'Kill the temp object (release it from memory)
Set ProcessEXE = Nothing

End Function

Public Function ListRunningProcesses()

' Function Get a list of all active Processes (tasks) running on your computer
' Based on the WMI (Windows Management Instrurmentation) code on MSDN
' This Function requires WMI (Windows Management Instrurmentation).
' WMI is part of Windows 2000, XP
' For more information see the MSDN Web site
' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/wmi_tasks__processes.asp
'
'This Function has no return value

Dim ProcessEXE            As Object
Dim x                     As Integer

x = 0

For Each ProcessEXE In GetObject("winmgmts://").InstancesOf("win32_process")
    'Go thru the list of processes and put into an array
    'We don't know how many array elements there are,
    'so redim the array "on the fly"
    
    ReDim Preserve ProcessArray(x + 1)
    ProcessArray(x) = ProcessEXE.Name
    x = x + 1
    
Next

ProcArraySize = x

'Kill the temp object (release it from memory)
Set ProcessEXE = Nothing

End Function
