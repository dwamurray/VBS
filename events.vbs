strComputer = "."

Set objSWbemServices = GetObject("winmgmts:\\" & strComputer)
Set colSWbemObjectSet = objSWbemServices.InstancesOf("Win32_NTLogEvent")
For Each objSWbemObject In colSWbemObjectSet    
    Wscript.Echo "Log File:        " & objSWbemObject.LogFile        & vbCrLf & _
                 "Record Number:   " & objSWbemObject.RecordNumber   & vbCrLf & _
                 "Type:            " & objSWbemObject.Type           & vbCrLf & _
                 "Time Generated:  " & objSWbemObject.TimeGenerated  & vbCrLf & _
                 "Source:          " & objSWbemObject.SourceName     & vbCrLf & _
                 "Category:        " & objSWbemObject.Category       & vbCrLf & _
                 "Category String: " & objSWbemObject.CategoryString & vbCrLf & _
                 "Event:           " & objSWbemObject.EventCode      & vbCrLf & _
                 "User:            " & objSWbemObject.User           & vbCrLf & _
                 "Computer:        " & objSWbemObject.ComputerName   & vbCrLf & _
                 "Message:         " & objSWbemObject.Message        & vbCrLf
Next
