Sub replace()

Dim myRange As Object
Set myRange = ActiveDocument.Content
myRange.Find.Execute FindText:="chupacabra", ReplaceWith:="hello", _
replace:=wdReplaceAll

End Sub