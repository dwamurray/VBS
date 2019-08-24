On Error Resume Next
Const CONVERSION_FACTOR = 1048576
Const WARNING_THRESHOLD = 100

If WScript.Arguments.Count = 0 Then
    Wscript.Echo "Usage: FirstScript.vbs server1 [server2] [server3] ..."
    WScript.Quit
End If

For Each Computer In WScript.Arguments
    Set objWMIService = GetObject("winmgmts://" & Computer)
    If Err.Number <> 0 Then
        Wscript.Echo Computer & " " & Err.Description
        Err.Clear
    Else
        Set colLogicalDisk = _
            objWMIService.InstancesOf("Win32_LogicalDisk")
        For Each objLogicalDisk In colLogicalDisk
            FreeMegaBytes = objLogicalDisk.FreeSpace / CONVERSION_FACTOR
            If FreeMegaBytes < WARNING_THRESHOLD Then
                Wscript.Echo Computer & " " & objLogicalDisk.DeviceID & _
                    " is low on disk space."
            End If
        Next
    End If
Next
