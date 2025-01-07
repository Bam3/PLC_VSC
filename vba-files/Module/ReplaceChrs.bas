Attribute VB_Name = "ReplaceChrs"
'novo dodano!

Sub pokaziFormo()
    UserFormSumniki.Show
End Sub

Sub popraviSumnike(ws As Worksheet)
    'Funkcija gre èez celotni dokument in popravi od proficyja zdrkane šumnike
    Dim findStrings As Variant
    Dim replaceStrings As Variant
    Dim i As Long

    ' Set up arrays with find and replace strings
    ' vèasih dobim ene vrste znakcov, vèasih pa drugaène. No idea why.
    findStrings = Array("Ä", "Å¡", "Å¾", "ÄŒ", "Å ", "Å½", "ÄT", "L~?", "Ll", "ÄS", "Š", "", "A¨", "A~?", "â€“", Â)
    replaceStrings = Array("è", "š", "", "È", "Š", "", "è", "š", "", "È", "Š", "", "è", "È", "-", "")

    ' Replace strings in the selected worksheet
    With ws.Cells
        For i = LBound(findStrings) To UBound(findStrings)
            .Replace what:=findStrings(i), replacement:=replaceStrings(i), LookAt:=xlPart, MatchCase:=True
        Next i
    End With
    
    MsgBox "Konèano!"
    
End Sub




