Const CONVERSION_FACTOR = 1048576
Const WARNING_THRESHOLD = 100
Computers = Array("atl-dc-01", "atl-dc-02", "atl-dc-03")
For Each Computer In Computers
    Set objWMIService = GetObject("winmgmts://" & Computer)
    Set colLogicalDisk = objWMIService.InstancesOf("Win32_LogicalDisk")
    For Each objLogicalDisk In colLogicalDisk
        FreeMegaBytes = objLogicalDisk.FreeSpace / CONVERSION_FACTOR
        If FreeMegaBytes < WARNING_THRESHOLD Then
            Wscript.Echo Computer & " " & objLogicalDisk.DeviceID & _
                " is low on disk space."
        End If
    Next
Next
