On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery ("Select * from Win32_Service ")
For Each objItem in colItems
    Wscript.Echo "DisplayName: " & objService.displayname
    Wscript.Echo "Name: " & objItem.name
    Wscript.Echo "State: " & objItem.state
    Wscript.Echo "StartMode: " & objItem.startmode
    Wscript.Echo "StartName: " & objItem.startname
    Wscript.Echo
Next