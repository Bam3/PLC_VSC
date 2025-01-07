Attribute VB_Name = "DopolniPraznaPoljaIOtabela"
Sub dopolniPraznaPoljaBtn_Click()
Dim endLoop As Boolean
Dim Row As Integer
Dim naslov As String
Row = 2
endLoop = False
    Do Until endLoop
        If Sheets("IOT").Cells(Row, "A").Value Like Empty Then
            endLoop = True
            MsgBox "Vstavljanje rezerv za DI, DO, AI in AO konèano!"
            Exit Sub
        End If
        If Sheets("IOT").Cells(Row, "B").Value Like Empty Or _
           Sheets("IOT").Cells(Row, "B").Value Like "/" Then
            naslov = Replace(Sheets("IOT").Cells(Row, "A").Value, "%", "")
            Sheets("IOT").Cells(Row, "B").Value = "REZ"
            Sheets("IOT").Cells(Row, "C").Value = naslov
            Sheets("IOT").Cells(Row, "D").Value = "Rezerva " & naslov
        End If
        Row = Row + 1
    Loop
End Sub
