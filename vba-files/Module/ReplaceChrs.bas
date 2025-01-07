Attribute VB_Name = "ReplaceChrs"
'novo dodano!

Sub pokaziFormo()
    UserFormSumniki.Show
End Sub

Sub popraviSumnike(ws As Worksheet)
    'Funkcija gre �ez celotni dokument in popravi od proficyja zdrkane �umnike
    Dim findStrings As Variant
    Dim replaceStrings As Variant
    Dim i As Long

    ' Set up arrays with find and replace strings
    ' v�asih dobim ene vrste znakcov, v�asih pa druga�ne. No idea why.
    findStrings = Array("Ä", "Å¡", "Å¾", "ÄŒ", "Å ", "Å½", "ÄT", "L~?", "Ll", "ÄS", "Š", "Ž", "A¨", "A~?", "â€“", Â,"È")
    replaceStrings = Array("č", "š", "ž", "č", "Š", "Ž", "č", "š", "ž", "Č", "Š", "Ž", "č", "Č", "-", "","Č")

    ' Replace strings in the selected worksheet
    With ws.Cells
        For i = LBound(findStrings) To UBound(findStrings)
            .Replace what:=findStrings(i), replacement:=replaceStrings(i), LookAt:=xlPart, MatchCase:=True
        Next i
    End With
    
    MsgBox "Končano!"
    
End Sub




