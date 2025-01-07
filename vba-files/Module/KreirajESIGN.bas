Attribute VB_Name = "KreirajESIGN"

Public Sub createESIGN()
    OsnovnaFormaESIGN.Show
End Sub
Public Function A_EN(ByVal AREA As Variant, ByVal Counter2 As Integer, sistem As String, sheetName As String, tipParametra As String) As Integer

Dim ImeTocke, ImeTockeForDescription, Description, odmik, ES_Type, Units, Priority, LC, UC, Security, podsistem, DescriptionPopravek As String
Dim Counter, Row, Lines As Integer
' izraèun za zanko, da vidi do kje mora šteti - pregleda cel TGD
Lines = Sheets(sheetName).Cells(1, "A").Value
Lines = Lines + 1
For Counter = 2 To Lines
ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
'pregled ali smo na rezervi, èe smo jo preskoèimo
If ImeTocke Like "*REZ*" Then GoTo line1

' doloèanje vseh parametrov
If ImeTocke Like "*_A_EN" Then
    ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
    Description = Sheets(sheetName).Cells(Counter, "C").Value
    DescriptionPopravek = Replace(Description, "enable", "vklop alarmiranja")
    odmik = Sheets(sheetName).Cells(Counter, "B").Value
    entry_type = "D"
    ES_Type = "0"
    Unit = ""
    Priority = "LOW"
    LC = "IZKLOP"
    UC = "VKLOP"
    Security = "TEHNIK"
    podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
    ' zapis v ESIGN
    Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
    Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
    Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
    Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
    Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & AREA(3) & ";" & tipParametra)
    Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & AREA(3) & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
    Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
    Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
    Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
    Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
    Sheets("ESIGN").Cells(Counter2, "K").Value = (DescriptionPopravek)
    Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
    Counter2 = Counter2 + 1

End If
line1: Next Counter

A_EN = Counter2
End Function
Public Function KVIT(ByVal AREA As Variant, ByVal Counter2 As Integer, sistem As String, sheetName As String, tipParametra As String) As Integer

Dim ImeTocke, ImeTockeForDescription, Description, odmik, ES_Type, Units, Priority, LC, UC, Security, DescriptionPopravek As String
Dim Counter, Row, Lines As Integer

' izraèun za zanko, da vidi do kje mora šteti - pregleda cel TGD
Lines = Sheets(sheetName).Cells(1, "A").Value
Lines = Lines + 1

For Counter = 2 To Lines
    ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
    'pregled ali smo na rezervi, èe smo jo preskoèimo
    If ImeTocke Like "*REZ*" Then GoTo line1
    If ImeTocke Like "*_KVIT" Then
        ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
        Description = Sheets(sheetName).Cells(Counter, "C").Value
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        entry_type = "R"
        ES_Type = "0"
        Unit = ""
        Priority = "LOW"
        LC = "/"
        UC = "KVIT"
        Security = "TEHNIK"
        podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
        ' zapis v ESIGN
        Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
        Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
        Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
        Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
        Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sistem & ";" & tipParametra)
        Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sistem & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
        Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
        Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
        Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
        Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
        Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
        Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
        Counter2 = Counter2 + 1
    End If

line1: Next Counter

KVIT = Counter2
End Function
Public Function ALM(ByVal AREA As Variant, ByVal Counter2 As Integer, sklop As String, sheetName As String, suffix As String, tipParametra As String) As Integer

Dim ImeTocke, ImeTockeForDescription, Description, odmik, ES_Type, Units, Priority, LC, UC, Security, Oznaka, celoIme As String
Dim Counter, Row, Lines, Lines_AO, Counter_AI, numberOfElements As Integer
'____________________________________________________________________________________________________________

'najdeš prvi %AI
Counter_AI = getNextRowNumber("IOT", "A", "%AI")
Oznaka = Sheets("IOT").Cells(Counter_AI, "C").Value

Do While Oznaka <> ""
    sistem = Sheets("IOT").Cells(Counter_AI, "B").Value
    celoIme = (sistem & "_VA_" & Oznaka & "_" & suffix)
    LC = decimalneVejiceEnote(Sheets("IOT").Cells(Counter_AI, "G").Value, Sheets("IOT").Cells(Counter_AI, "E").Value)
    UC = decimalneVejiceEnote(Sheets("IOT").Cells(Counter_AI, "G").Value, Sheets("IOT").Cells(Counter_AI, "F").Value)
    Unit = Sheets("IOT").Cells(Counter_AI, "G").Value
    ' izraèun za zanko, da vidi do kje mora šteti - pregleda cel TGD
    Lines = Sheets(sheetName).Cells(1, "A").Value
    Lines = Lines + 1
    
    For Counter = 2 To Lines
        ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
        'pregled ali smo na rezervi, èe smo jo preskoèimo
        If ImeTocke Like "*REZ*" Then GoTo line1
                
        If ImeTocke Like celoIme Then
            entry_type = "A"
            ES_Type = "0"
            Priority = "LOW"
            Security = "TEHNOLOG"
            podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
            Description = Sheets(sheetName).Cells(Counter, "C").Value
            If suffix Like "HIHI" Then DescriptionPopravek = Replace(Description, "meja", "zg. alm. meja")
            If suffix Like "HI" Then DescriptionPopravek = Replace(Description, "meja", "zg. opo. meja")
            If suffix Like "LO" Then DescriptionPopravek = Replace(Description, "meja", "sp. opo. meja")
            If suffix Like "LOLO" Then DescriptionPopravek = Replace(Description, "meja", "sp. alm. meja")
            odmik = Sheets(sheetName).Cells(Counter, "B").Value
            'zapis v ESIGN
            Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
            Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
            Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
            Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
            Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sklop & ";" & tipParametra)
            Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sklop & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
            Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
            Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
            Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
            Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
            Sheets("ESIGN").Cells(Counter2, "K").Value = (DescriptionPopravek)
            Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
            Counter2 = Counter2 + 1
            GoTo line2 'zakljuèi in poišèi nov AI
        End If
line1: Next Counter
line2:
        Counter_AI = Counter_AI + 1
        Oznaka = Sheets("IOT").Cells(Counter_AI, "C").Value
Loop
ALM = Counter2
End Function
Public Function ZAK_AI(ByVal AREA As Variant, ByVal Counter2 As Integer, sistem As String, sheetName As String, suffix As String, tipParametra As String) As Integer

Dim ImeTocke, ImeTockeForDescription, Description, odmik, ES_Type, Units, Priority, LC, UC, Security As String
Dim Counter, Row, Lines As Integer

' izraèun za zanko, da vidi do kje mora šteti - pregleda cel TGD
Lines = Sheets(sheetName).Cells(1, "A").Value
Lines = Lines + 1

For Counter = 2 To Lines
ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
'pregled ali smo na rezervi, èe smo jo preskoèimo
If ImeTocke Like "*REZ*" Then GoTo line1
If ImeTocke Like "*" & suffix Then
    Description = Sheets(sheetName).Cells(Counter, "C").Value
    odmik = Sheets(sheetName).Cells(Counter, "B").Value
    entry_type = "A"
    ES_Type = "0"
    LC = "1"
    UC = "32000"
    Unit = "s"
    Priority = "LOW"
    Security = "TEHNOLOG"
    podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
    ' zapis v ESIGN
    Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
    Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
    Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
    Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
    Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sistem & ";" & tipParametra)
    Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sistem & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
    Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
    Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
    Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
    Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
    Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
    Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
    
    Counter2 = Counter2 + 1
End If
line1: Next Counter
ZAK_AI = Counter2
End Function
Public Function PID_ROCNO(ByVal AREA As Variant, ByVal Counter2 As Integer, sistem As String, sheetName As String, tipParametra As String) As Integer

Dim ImeTocke, ImeTockeForDescription, Description, odmik, ES_Type, Units, Priority, LC, UC, Security, DescriptionPopravek As String
Dim Counter, Row, Lines As Integer

' izraèun za zanko, da vidi do kje mora šteti - pregleda cel TGD
Lines = Sheets(sheetName).Cells(1, "A").Value
Lines = Lines + 1

For Counter = 2 To Lines
    ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
    'pregled ali smo na rezervi, èe smo jo preskoèimo
    If ImeTocke Like "*REZ*" Then GoTo line1
   
        If ImeTocke Like "*_PID*ROC" Then
            ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
            Description = Sheets(sheetName).Cells(Counter, "C").Value
            If Description Like "*rezerva*" Then GoTo line1
            odmik = Sheets(sheetName).Cells(Counter, "B").Value
            entry_type = "D"
            ES_Type = "0"
            Unit = ""
            Priority = "LOW"
            LC = "AVTO"
            UC = "ROÈNO"
            Security = "TEHNIK"
            podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
            If ImeTocke Like "*_SP_ROC" Then
                LC = "IZKLOP"
                UC = "VKLOP"
            End If
            If ImeTocke Like "*_VD_*" Then
                Prostor = Left(ImeTocke, InStr(ImeTocke, "_VD_") - 1)
            Else
                Prostor = Left(ImeTocke, InStr(ImeTocke, "_VA_") - 1)
            End If
            ' zapis v ESIGN
            Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
            Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
            Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
            Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
            Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sistem & ";" & tipParametra)
            Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sistem & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
            Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
            Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
            Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
            Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
            Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
            Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
            Counter2 = Counter2 + 1
        End If

line1: Next Counter

PID_ROCNO = Counter2
End Function
Public Function RAMP(ByVal AREA As Variant, ByVal Counter2 As Integer, sistem As String, sheetName As String, tipParametra As String) As Integer

Dim ImeTocke, ImeTockeForDescription, Description, odmik, ES_Type, Units, Priority, LC, UC, Security, DescriptionPopravek As String
Dim Counter, Row, Lines As Integer

' izraèun za zanko, da vidi do kje mora šteti - pregleda cel TGD
Lines = Sheets(sheetName).Cells(1, "A").Value
Lines = Lines + 1
For Counter = 2 To Lines
    ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
    'pregled ali smo na rezervi, èe smo jo preskoèimo
    If ImeTocke Like "REZ*" Then GoTo line1
    ' doloèanje vseh parametrov
    If ImeTocke Like "*_VA_RAMP*" Then
        ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
        Description = Sheets(sheetName).Cells(Counter, "C").Value
        If Description Like "*rezerva*" Then GoTo line1
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        entry_type = "A"
        ES_Type = "0"
        Unit = "%"
        Priority = "LOW"
        Security = "TEHNOLOG"
        LC = "0"
        UC = "100"
        podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
        ' zapis v ESIGN
        Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
        Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
        Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
        Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
        Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sistem & ";" & tipParametra)
        Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sistem & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
        Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
        Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
        Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
        Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
        Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
        Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
        Counter2 = Counter2 + 1
    End If
line1: Next Counter
RAMP = Counter2
End Function

Public Function VA_PID(ByVal AREA As Variant, ByVal Counter2 As Integer, sistem As String, sheetName As String, tipParametra As String) As Integer

Dim ImeTocke, ImeTockeForDescription, Description, odmik, ES_Type, Units, Priority, LC, UC, Security, DescriptionPopravek, WrdArray(), Ime_SP_SC, Oznaka As String
Dim Counter, Row, Lines As Integer


' izraèun za zanko, da vidi do kje mora šteti - pregleda cel TGD
Lines = Sheets(sheetName).Cells(1, "A").Value
Lines = Lines + 1
For Counter = 2 To Lines
    ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
    Description = Sheets(sheetName).Cells(Counter, "C").Value
    
    'pregled ali smo na rezervi, èe smo jo preskoèimo
    If ImeTocke Like "*REZ*" Or Description Like "*rezerva*" Then GoTo line1
    sistem = OsnovnaFormaESIGN.TextBox_sistem.Value
    'doloèanje vseh parametrov
    If ImeTocke Like "*_PID*KP" Then
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        entry_type = "A"
        ES_Type = "0"
        Unit = ""
        Priority = "LOW"
        LC = "-10,00"
        UC = "10,00"
        Security = "TEHNOLOG"
        podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
        ' zapis v ESIGN
        Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
        Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
        Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
        Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
        Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sistem & ";" & tipParametra)
        Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sistem & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
        Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
        Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
        Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
        Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
        Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
        Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
        Counter2 = Counter2 + 1
    End If
    
    If ImeTocke Like "*_PID*KI" Then
        ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
        Description = Sheets(sheetName).Cells(Counter, "C").Value
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        entry_type = "A"
        ES_Type = "0"
        Unit = ""
        Priority = "LOW"
        LC = "0,00"
        UC = "10,00"
        Security = "TEHNOLOG"
        podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
        Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
        Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
        Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
        Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
        Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sistem & ";" & tipParametra)
        Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sistem & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
        Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
        Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
        Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
        Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
        Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
        Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
        Counter2 = Counter2 + 1
    End If
    
    If ImeTocke Like "*_PID*_UC" Then
        ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
        Description = Sheets(sheetName).Cells(Counter, "C").Value
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        entry_type = "A"
        ES_Type = "0"
        Priority = "LOW"
        Security = "TEHNOLOG"
        LC = "0,0"
        UC = "100,0"
        Unit = "%"
        podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
        ' zapis v ESIGN
        Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
        Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
        Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
        Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
        Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sistem & ";" & tipParametra)
        Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sistem & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
        Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
        Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
        Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
        Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
        Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
        Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
        Counter2 = Counter2 + 1
    End If
    
    If ImeTocke Like "*_PID*_LC" Then
        ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
        Description = Sheets(sheetName).Cells(Counter, "C").Value
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        entry_type = "A"
        ES_Type = "0"
        Priority = "LOW"
        Security = "TEHNOLOG"
        LC = "0,0"
        UC = "100,0"
        Unit = "%"
        podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
        ' zapis v ESIGN
        Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
        Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
        Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
        Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
        Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sistem & ";" & tipParametra)
        Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sistem & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
        Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
        Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
        Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
        Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
        Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
        Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
        Counter2 = Counter2 + 1
    End If
    
    If ImeTocke Like "*_PID*_CV" Then
        ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
        Description = Sheets(sheetName).Cells(Counter, "C").Value
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        entry_type = "A"
        ES_Type = "0"
        Priority = "LOW"
        Security = "TEHNIK"
        LC = "0,0"
        UC = "100,0"
        Unit = "%"
        podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
        ' zapis v ESIGN
        Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
        Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
        Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
        Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
        Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sistem & ";" & tipParametra)
        Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sistem & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
        Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
        Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
        Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
        Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
        Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
        Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
        Counter2 = Counter2 + 1
    End If
    
    If ImeTocke Like "*_PID*_OPV" Then
        ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
        Description = Sheets(sheetName).Cells(Counter, "C").Value
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        entry_type = "A"
        ES_Type = "0"
        Priority = "LOW"
        Security = "TEHNOLOG"
        LC = "0,0"
        UC = "100,0"
        Unit = "%"
        podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
        ' zapis v ESIGN
        Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
        Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
        Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
        Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
        Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sistem & ";" & tipParametra)
        Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sistem & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
        Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
        Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
        Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
        Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
        Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
        Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
        Counter2 = Counter2 + 1
    End If
    
    
    If ImeTocke Like "*_PID*_SP" Or ImeTocke Like "*_PID*_SP_DEJ" Then
        ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
        Description = Sheets(sheetName).Cells(Counter, "C").Value
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        entry_type = "A"
        ES_Type = "0"
        Priority = "LOW"
        Security = "TEHNOLOG"
        podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
        If Description Like "*tlak*" Then
            Unit = "bar"
        ElseIf Description Like "*Reg. temp*" Then
            Unit = "°C"
        Else
            Unit = "??"
        End If
        'da bomo to lazje najdli kasneje
        LC = "xxx"
        UC = "xxx"
        
        ' zapis v Sheet2
        Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
        Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
        Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
        Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
        Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sistem & ";" & tipParametra)
        Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sistem & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
        Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
        Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
        Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
        If LC Like "xxx" Then Sheets("ESIGN").Cells(Counter2, "I").Interior.ColorIndex = 3
        Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
        If UC Like "xxx" Then Sheets("ESIGN").Cells(Counter2, "J").Interior.ColorIndex = 3
        Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
        Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
        'Sheets("ESIGN").range("A"
        Counter2 = Counter2 + 1
    End If
    
    
    ' Ta if stavek SP èlenu kanalskega tlaka zamenja odmik z odmikom skaliranega kanalskega tlaka (SP_SC)
    If ImeTocke Like "*_PID*_SP_SC" Then
        Ime_SP_SC = Left(ImeTocke, Len(ImeTocke) - 3)
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        Counter_AO = 1
        Oznaka = Sheets("ESIGN").Cells(Counter_AO, "A").Value
        Do While Oznaka <> ""
            If Oznaka Like Ime_SP_SC Then Sheets("ESIGN").Cells(Counter_AO, "B").Value = (odmik)
            Counter_AO = Counter_AO + 1
            Oznaka = Sheets("ESIGN").Cells(Counter_AO, "A").Value
        Loop

    End If

line1: Next Counter

VA_PID = Counter2
End Function

Public Function AO(ByVal AREA As Variant, ByVal Counter2 As Integer, sklop As String, sheetName As String, tipParametra As String) As Integer

Dim ImeTocke, ImeTockeForDescription, Description, odmik, ES_Type, Units, Priority, LC, UC, Security, Oznaka, celoIme As String
Dim Counter, Row, Lines, Lines_AO, Counter_AO As Integer

'najdeš prvi %AI
Counter_AO = getNextRowNumber("IOT", "A", "%AQ")
Oznaka = Sheets("IOT").Cells(Counter_AO, "C").Value

Do While Not Oznaka Like "AQ*"
    podsistem = Sheets("IOT").Cells(Counter_AO, "B").Value
    celoIme = (sistem & "_VA_" & Oznaka)
    LC = decimalneVejiceEnote(Sheets("IOT").Cells(Counter_AO, "G").Value, Sheets("IOT").Cells(Counter_AO, "E").Value)
    UC = decimalneVejiceEnote(Sheets("IOT").Cells(Counter_AO, "G").Value, Sheets("IOT").Cells(Counter_AO, "F").Value)
    Unit = Sheets("IOT").Cells(Counter_AO, "G").Value
    ' izraèun za zanko, da vidi do kje mora šteti - pregleda cel TGD
    Lines = Sheets(sheetName).Cells(1, "A").Value
    Lines = Lines + 1
    
    For Counter = 2 To Lines
        ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
        'pregled ali smo na rezervi, èe smo jo preskoèimo
        If ImeTocke Like "*REZ*" Then GoTo line1
        ' doloèanje vseh parametrov
        If ImeTocke Like celoIme Then
        ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
        Description = Sheets(sheetName).Cells(Counter, "C").Value
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        entry_type = "A"
        ES_Type = "0"
        Priority = "LOW"
        Security = "TEHNIK"
        podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
        ' zapis v ESIGN
        Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
        Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
        Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
        Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
        Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sklop & ";" & tipParametra)
        Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sklop & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
        Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
        Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
        Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
        Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
        Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
        Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
        Counter2 = Counter2 + 1
        GoTo line2
        End If
line1:
        Next Counter
line2:
    Counter_AO = Counter_AO + 1
    Oznaka = Sheets("IOT").Cells(Counter_AO, "C").Value
Loop
AO = Counter2
End Function

Public Function SCADA_ESIGN(ByVal AREA As Variant, ByVal Counter2 As Integer, sklop As String, sheetName As String, suffix As String, tipParametra As String) As Integer

Dim ImeTocke, ImeTockeForDescription, Description, odmik, ES_Type, Units, Priority, LC, UC, Security, podsistem, DescriptionPopravek As String
Dim Counter, Row, Lines As Integer
' izraèun za zanko, da vidi do kje mora šteti - pregleda cel TGD
Lines = Sheets(sheetName).Cells(1, "A").Value
Lines = Lines + 1
For Counter = 2 To Lines
    ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
    'pregled ali smo na rezervi, èe smo jo preskoèimo
    If ImeTocke Like "*SIS*" Then GoTo line1
    
    ' doloèanje vseh parametrov
    If ImeTocke Like "*_" & suffix Then
        ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
        Description = Sheets(sheetName).Cells(Counter, "C").Value
        If suffix Like "VKLOP_SCADA" Then DescriptionPopravek = Replace(Description, "Vklop Scada", "Vklop sistema")
        If suffix Like "KVIT_SCADA" Then DescriptionPopravek = Replace(Description, "Kvit Scada", "Reset sistema")
        If suffix Like "AUTO" Then DescriptionPopravek = Replace(Description, "Avtomatsko", "Vklop avtomatskega vodenja")
        If suffix Like "ROCNO" Then DescriptionPopravek = Replace(Description, "Kvit Scada", "Vklop roènega vodenja")
        If suffix Like "SERVIS" Then DescriptionPopravek = Replace(Description, "Kvit Scada", "Vklop servisnega vodenja")
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        If suffix Like "VKLOP_SCADA" Then entry_type = "D" Else entry_type = "R"
        ES_Type = "0"
        Unit = ""
        Priority = "LOW"
        If suffix Like "VKLOP_SCADA" Then LC = "IZKLOP" Else LC = "/"
        If suffix Like "VKLOP_SCADA" Then
            UC = "VKLOP"
        ElseIf suffix Like "KVIT_SCADA" Then
            UC = "KVIT"
        Else
            UC = suffix
        End If
        If suffix Like "SERVIS" Then Security = "TEHNOLOG" Else Security = "TEHNIK"
        podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
        ' zapis v ESIGN
        Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
        Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
        Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
        Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
        Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sklop & ";" & tipParametra)
        Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sklop & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
        Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
        Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
        Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
        Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
        Sheets("ESIGN").Cells(Counter2, "K").Value = (DescriptionPopravek)
        Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
        Counter2 = Counter2 + 1

End If
line1: Next Counter

SCADA_ESIGN = Counter2
End Function

Public Function OBRH_ST_VKL(ByVal AREA As Variant, ByVal Counter2 As Integer, sistem As String, sheetName As String, tipParametra As String) As Integer

Dim ImeTocke, ImeTockeForDescription, Description, odmik, ES_Type, Units, Priority, LC, UC, Security, podsistem, DescriptionPopravek As String
Dim Counter, Row, Lines As Integer
' izraèun za zanko, da vidi do kje mora šteti - pregleda cel TGD
Lines = Sheets(sheetName).Cells(1, "A").Value
Lines = Lines + 1
For Counter = 2 To Lines
    ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
    'pregled ali smo na rezervi, èe smo jo preskoèimo
    If ImeTocke Like "*SIS*" Then GoTo line1
    
    ' doloèanje vseh parametrov
    If ImeTocke Like "*OBRHD*" Or ImeTocke Like "*ST_VKL*" Then
        ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
        Description = Sheets(sheetName).Cells(Counter, "C").Value
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        If Sheets(sheetName).Cells(Counter, "A").Value Like "*_VA_*" Then entry_type = "A" Else entry_type = "R"
        ES_Type = "0"
        If Sheets(sheetName).Cells(Counter, "A").Value Like "*_VA_*" And ImeTocke Like "*_OBRHD" Then Unit = "h" Else Unit = ""
        
        Priority = "LOW"
        If entry_type Like "A" Then LC = "0" Else LC = "/"
        If entry_type Like "A" Then UC = "999999" Else UC = "RESET"
        Security = "TEHNOLOG"
        podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
        ' zapis v ESIGN
        Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
        Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
        Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
        Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
        Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sistem & ";" & tipParametra)
        Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sistem & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
        Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
        Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
        Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
        Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
        Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
        Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
        Counter2 = Counter2 + 1

End If
line1: Next Counter

OBRH_ST_VKL = Counter2
End Function
Public Function DI_SRV(ByVal AREA As Variant, ByVal Counter2 As Integer, sklop As String, sheetName As String, suffix As String, tipParametra As String) As Integer

Dim ImeTocke, ImeTockeForDescription, Description, odmik, ES_Type, Units, Priority, LC, UC, Security, podsistem, DescriptionPopravek As String
Dim Counter, Row, Lines As Integer
Dim Counter_DI As Integer

'najdeš prvi %AI
Counter_DI = getNextRowNumber("IOT", "A", "%I")
Oznaka = Sheets("IOT").Cells(Counter_DI, "C").Value
Do While Oznaka <> ""
    sistem = Sheets("IOT").Cells(Counter_DI, "B").Value
    celoIme = (sistem & "_VD_" & Oznaka & "_" & suffix)
    If suffix Like "SV" Then
        LC = Sheets("IOT").Cells(Counter_DI, "E").Value
        UC = Sheets("IOT").Cells(Counter_DI, "F").Value
    Else
        LC = "AUTO"
        UC = "SERVIS"
    End If
    Unit = ""
    ' izraèun za zanko, da vidi do kje mora šteti - pregleda cel TGD
    Lines = Sheets(sheetName).Cells(1, "A").Value
    Lines = Lines + 1
    
    For Counter = 2 To Lines
        ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
        'pregled ali smo na rezervi, èe smo jo preskoèimo
        If ImeTocke Like "*REZ*" Then GoTo line1
                
        If ImeTocke Like celoIme Then
            entry_type = "D"
            ES_Type = "0"
            Priority = "LOW"
            Security = "TEHNOLOG"
            podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
            Description = Sheets(sheetName).Cells(Counter, "C").Value
            odmik = Sheets(sheetName).Cells(Counter, "B").Value
            'zapis v ESIGN
            Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
            Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
            Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
            Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
            Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sklop & ";" & tipParametra)
            Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sklop & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
            Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
            Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
            Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
            Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
            Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
            Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
            Counter2 = Counter2 + 1
            GoTo line2 'zakljuèi in poišèi nov AI
        End If
line1: Next Counter
line2:
        Counter_DI = Counter_DI + 1
        Oznaka = Sheets("IOT").Cells(Counter_DI, "C").Value
Loop
DI_SRV = Counter2
End Function

Public Function REZ_ACT(ByVal AREA As Variant, ByVal Counter2 As Integer, sklop As String, sheetName As String, tipParametra As String) As Integer

Dim ImeTocke, ImeTockeForDescription, Description, odmik, ES_Type, Units, Priority, LC, UC, Security, podsistem, DescriptionPopravek As String
Dim Counter, Row, Lines As Integer
' izraèun za zanko, da vidi do kje mora šteti - pregleda cel TGD
Lines = Sheets(sheetName).Cells(1, "A").Value
Lines = Lines + 1
For Counter = 2 To Lines
    ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
    'pregled ali smo na rezervi, èe smo jo preskoèimo
    If ImeTocke Like "*REZ*" Then GoTo line1
    
    ' doloèanje vseh parametrov
    If ImeTocke Like "*_RZ" Then
        ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
        Description = Sheets(sheetName).Cells(Counter, "C").Value
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        entry_type = "E"
        ES_Type = "0"
        Unit = ""
        Priority = "LOW"
        LC = "0"
        UC = "3"
        Security = "TEHNIK"
        podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
        ' zapis v ESIGN
        Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
        Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
        Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
        Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
        Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sklop & ";" & tipParametra)
        Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sklop & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
        Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
        Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
        Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
        Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
        Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
        Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
        Counter2 = Counter2 + 1
        Call createTABforRZ(odmik, UC)
End If
line1: Next Counter

REZ_ACT = Counter2
End Function
Public Function DI_MAN_SRV(ByVal AREA As Variant, ByVal Counter2 As Integer, sklop As String, sheetName As String, suffix As String, tipParametra As String) As Integer

Dim ImeTocke, ImeTockeForDescription, Description, odmik, ES_Type, Units, Priority, LC, UC, Security, podsistem, DescriptionPopravek As String
Dim Counter, Row, Lines As Integer
Dim Counter_DI As Integer

' izraèun za zanko, da vidi do kje mora šteti - pregleda cel TGD
Lines = Sheets(sheetName).Cells(1, "A").Value
Lines = Lines + 1

For Counter = 2 To Lines
    ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
    'pregled ali smo na rezervi, èe smo jo preskoèimo
    If ImeTocke Like "*REZ*" Then GoTo line1
            
    If ImeTocke Like "*_VD_*_" & suffix Then
        entry_type = "D"
        ES_Type = "0"
        Priority = "LOW"
        If suffix Like "SR" Then Security = "TEHNOLOG" Else Security = "TEHNIK"
        podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
        If Description Like "*ventil *" Or Description Like "*Ventil *" Then
            LC = "ZAPRI"
            UC = "ODPRI"
        Else
            LC = "IZKLOPI"
            UC = "VKLOPI"
        End If
        Description = Sheets(sheetName).Cells(Counter, "C").Value
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        'zapis v ESIGN
        Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
        Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
        Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
        Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
        Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sklop & ";" & tipParametra)
        Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sklop & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
        Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
        Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
        Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
        Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
        Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
        Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
        Counter2 = Counter2 + 1
    End If
line1: Next Counter

DI_MAN_SRV = Counter2
End Function
Public Function VA_MAN_SRV(ByVal AREA As Variant, ByVal Counter2 As Integer, sklop As String, sheetName As String, suffix As String, tipParametra As String) As Integer

Dim ImeTocke, ImeTockeForDescription, Description, odmik, ES_Type, Units, Priority, LC, UC, Security, podsistem, DescriptionPopravek As String
Dim Counter, Row, Lines As Integer
Dim Counter_DI As Integer

' izraèun za zanko, da vidi do kje mora šteti - pregleda cel TGD
Lines = Sheets(sheetName).Cells(1, "A").Value
Lines = Lines + 1

For Counter = 2 To Lines
    ImeTocke = Sheets(sheetName).Cells(Counter, "A").Value
    'pregled ali smo na rezervi, èe smo jo preskoèimo
    If ImeTocke Like "*REZ*" Then GoTo line1
            
    If ImeTocke Like "*_VA_*_" & suffix Then
        entry_type = "A"
        ES_Type = "0"
        Priority = "LOW"
        If suffix Like "SR" Then Security = "TEHNOLOG" Else Security = "TEHNIK"
        podsistem = Left(ImeTocke, InStr(ImeTocke, "_") - 1)
        LC = "0,0"
        UC = "100,0"
        Description = Sheets(sheetName).Cells(Counter, "C").Value
        odmik = Sheets(sheetName).Cells(Counter, "B").Value
        Unit = "%"
        'zapis v ESIGN
        Sheets("ESIGN").Cells(Counter2, "A").Value = (ImeTocke)
        Sheets("ESIGN").Cells(Counter2, "B").Value = (odmik)
        Sheets("ESIGN").Cells(Counter2, "C").Value = (ES_Type)
        Sheets("ESIGN").Cells(Counter2, "D").Value = (Security)
        Sheets("ESIGN").Cells(Counter2, "E").Value = (podsistem & ";" & sklop & ";" & tipParametra)
        Sheets("ESIGN").Cells(Counter2, "F").Value = (AREA(0) & ";" & AREA(1) & ";" & podsistem & ";" & sklop & ";" & AREA(4) & ";" & tipParametra & ";" & AREA(6) & ";" & AREA(7) & ";" & AREA(8) & ";" & AREA(9) & ";")
        Sheets("ESIGN").Cells(Counter2, "G").Value = (Priority)
        Sheets("ESIGN").Cells(Counter2, "H").Value = (entry_type)
        Sheets("ESIGN").Cells(Counter2, "I").Value = (LC)
        Sheets("ESIGN").Cells(Counter2, "J").Value = (UC)
        Sheets("ESIGN").Cells(Counter2, "K").Value = (Description)
        Sheets("ESIGN").Cells(Counter2, "L").Value = (Unit)
        Counter2 = Counter2 + 1
    End If
line1: Next Counter

VA_MAN_SRV = Counter2
End Function


Public Function createTABforRZ(ByVal odmik As String, ByVal UC As String)

If Sheets("ESIGN_TAB").Cells(1, "A").Value Like "" Then
    Counter = 1
Else
    Counter = getNextRowNumber("ESIGN_TAB", "A", "")
End If

Sheets("ESIGN_TAB").Cells(Counter, "A").Value = odmik
Sheets("ESIGN_TAB").Cells(Counter, "B").Value = UC

Sheets("ESIGN_TAB").Cells(Counter + 1, "A").Value = "1"
Sheets("ESIGN_TAB").Cells(Counter + 1, "B").Value = "AVTO"
Sheets("ESIGN_TAB").Cells(Counter + 1, "C").Value = "TEHNIK"

Sheets("ESIGN_TAB").Cells(Counter + 2, "A").Value = "2"
Sheets("ESIGN_TAB").Cells(Counter + 2, "B").Value = "ROÈNO"
Sheets("ESIGN_TAB").Cells(Counter + 2, "C").Value = "TEHNIK"

Sheets("ESIGN_TAB").Cells(Counter + 3, "A").Value = "3"
Sheets("ESIGN_TAB").Cells(Counter + 3, "B").Value = "SERVIS"
Sheets("ESIGN_TAB").Cells(Counter + 3, "C").Value = "TEHNOLOG"

End Function












