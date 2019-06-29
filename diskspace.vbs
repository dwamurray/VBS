Const CONVERSION_FACTOR = 1048576 
Set objWMIService = GetObject("winmgmts:") 
Set objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='c:'") 
FreeMegaBytes = objLogicalDisk.FreeSpace / CONVERSION_FACTOR 
Wscript.Echo "There are " & Int(FreeMegaBytes) & " megabytes free on C:"