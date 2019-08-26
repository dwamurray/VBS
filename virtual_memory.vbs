strComputer = "."

Set objSWbemServices = GetObject("winmgmts:\\" & strComputer)
Set colSWbemObjectSet = _
    objSWbemServices.InstancesOf("Win32_LogicalMemoryConfiguration")

For Each objSWbemObject In colSWbemObjectSet
    Wscript.Echo "Total Virtual Memory (kb): " & _
        objSWbemObject.TotalVirtualMemory
Next
