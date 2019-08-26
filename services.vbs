strComputer = "."

Set objSWbemServices = GetObject("winmgmts:\\" & strComputer)
Set colSWbemObjectSet = objSWbemServices.InstancesOf("Win32_Service")

For Each objSWbemObject In colSWbemObjectSet
    Wscript.Echo "Display Name:  " & objSWbemObject.DisplayName & vbCrLf & _
                 "   State:      " & objSWbemObject.State       & vbCrLf & _
                 "   Start Mode: " & objSWbemObject.StartMode
Next
