Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile("c:\windows\system32\scrrun.dll")
Wscript.Echo "Date created: " & objFile.DateCreated
Wscript.Echo "Date last accessed: " & objFile.DateLastAccessed
Wscript.Echo "Date last modified: " & objFile.DateLastModified
Wscript.Echo "Drive: " & objFile.Drive
Wscript.Echo "Name: " & objFile.Name
Wscript.Echo "Parent folder: " & objFile.ParentFolder
Wscript.Echo "Path: " & objFile.Path
Wscript.Echo "Short name: " & objFile.ShortName
Wscript.Echo "Short path: " & objFile.ShortPath
Wscript.Echo "Size: " & objFile.Size
Wscript.Echo "Type: " & objFile.Type

