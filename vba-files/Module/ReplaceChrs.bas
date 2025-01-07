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
    findStrings = Array("č", "š", "ž", "Č", "� ", "Ž", "�T", "L~?", "Ll", "�S", "�", "�", "A�", "A~?", "–", �)
    replaceStrings = Array("�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "-", "")

    ' Replace strings in the selected worksheet
    With ws.Cells
        For i = LBound(findStrings) To UBound(findStrings)
            .Replace what:=findStrings(i), replacement:=replaceStrings(i), LookAt:=xlPart, MatchCase:=True
        Next i
    End With
    
    MsgBox "Kon�ano!"
    
End Sub




