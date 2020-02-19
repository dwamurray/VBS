strComputer = "."

Set objSWbemServices = GetObject("winmgmts:\\" & strComputer)
Set colOperatingSystems = objSWbemServices.InstancesOf("Win32_OperatingSystem")

For Each objOperatingSystem In colOperatingSystems
    Wscript.Echo "Name: " & objOperatingSystem.Name   & vbCrLf & _
    "Caption: " & objOperatingSystem.Caption         & vbCrLf & _
    "CurrentTimeZone: " & objOperatingSystem.CurrentTimeZone & vbCrLf & _
    "LastBootUpTime: " & objOperatingSystem.LastBootUpTime  & vbCrLf & _
    "LocalDateTime: " & objOperatingSystem.LocalDateTime   & vbCrLf & _
    "Locale: " & objOperatingSystem.Locale          & vbCrLf & _
    "Manufacturer: " & objOperatingSystem.Manufacturer    & vbCrLf & _
    "OSType: " & objOperatingSystem. OSType         & vbCrLf & _
    "Version: " & objOperatingSystem.Version         & vbCrLf & _
    "Service Pack: " & objOperatingSystem.ServicePackMajorVersion  & _
    "." & objOperatingSystem.ServicePackMinorVersion           & vbCrLf & _
    "Windows Directory: " & objOperatingSystem.WindowsDirectory
Next
