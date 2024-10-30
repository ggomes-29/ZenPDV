Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.CopyFile WScript.Arguments.Item(0), WScript.Arguments.Item(1), True
