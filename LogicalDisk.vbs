strComputer = "."

Set objSWbemServices = GetObject("winmgmts:\\" & strComputer)
Set colSWbemObjectSet = objSWbemServices.InstancesOf("Win32_LogicalDisk")

For Each objSWbemObject In colSWbemObjectSet
    Wscript.Echo objSWbemObject.DeviceID
Next
