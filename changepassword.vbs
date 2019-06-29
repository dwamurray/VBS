Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("server.txt", ForReading)
 
Const ForReading = 1
strUser = "dooda"
Dim arrFileLines()
i = 0
Do Until objFile.AtEndOfStream
Redim Preserve arrFileLines(i)
arrFileLines(i) = objFile.ReadLine
i = i + 1
Loop
objFile.Close
 
 
 
 
On Error Resume Next
ErrorOccurred = False
 
For Each strLine in arrFileLines
 
strComputer = strLine
SET oFS = CreateObject("Scripting.FileSystemObject")
 
strPath = "C:\Documents and Settings\murrayd\Desktop\"
strFileName = strLine & "-Log" & ".txt"
strFullName = objFSO.BuildPath(strPath, strFileName)
Set objFile = objFSO.CreateTextFile(strFullName)
objFile.Close
 
SET Gfile = oFS.GetFile(strFullName)
SET Wfile = Gfile.OpenAsTextStream(8,-1)
 
Wfile.WriteLine "Resetting Password for: " &strComputer
 
SET objUser = GETOBJECT("WinNT://" & strLine & "/" & strUser)
objUser.SetPassword "!B0nanz@3389"
objUser.SetInfo
 
if err.number <> 0 then
Wfile.WriteLine "Error connecting to " &strComputer
err.clear
ErrorOccurred = True
else
Wfile.WriteLine "Password set for " &objUser.name
end if
               
Wfile.Close
Next
WSCript.Quit