Const CONVERSION_FACTOR = 1048576
Const WARNING_THRESHOLD = 100
Computer = "dc01"
Set objWMIService = GetObject("winmgmts://" & Computer)
Set colLogicalDisk = objWMIService.InstancesOf("Win32_LogicalDisk")
For Each objLogicalDisk In colLogicalDisk
    FreeMegaBytes = objLogicalDisk.FreeSpace / CONVERSION_FACTOR
    If FreeMegaBytes < WARNING_THRESHOLD Then
        Wscript.Echo objLogicalDisk.DeviceID & " is low on disk space."
    Else
       Wscript.Echo objLogicalDisk.DeviceID & " has adequate disk space."
    End If
Next
