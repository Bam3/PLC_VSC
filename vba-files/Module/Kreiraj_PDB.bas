Attribute VB_Name = "Kreiraj_PDB"
Sub PDB_Create()
Dim codeTypes As Object
Set codeTypes = CreateObject("Scripting.Dictionary")
Dim numberOfElements As Long
Dim Description As String
Dim system As String
Dim tag As String
Dim i As Long
Dim msg As Boolean
Dim address As String
Dim PLCName As String
Dim SCADAname As String
Dim Registri() As String
Dim registriUNI() As Variant
Dim parsed() As String
Dim parsedNext() As String
Dim EndLine As String
EndLine = "[-------------------------------------------------End of Block List-------------------------------------------------]"

    'we delete and then create all the sheets
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("PDB").Delete
    Sheets.Add.Name = "PDB"
    PLCName = Sheets("IOT").Cells(1, "I").Value
    SCADAname = Sheets("IOT").Cells(2, "I").Value
    
    ' oèisti TGD substitutione tako da dobiš samo registre
    numberOfElements = Sheets("TGD").Cells(1, "A").Value
    ReDim Registri(numberOfElements)
    For i = LBound(Registri) To UBound(Registri)
        parsed = Split(Sheets("TGD").Cells(i, "B").Value, ".")
        For j = LBound(parsed) To UBound(parsed)
            If parsed(j) Like "*AR*" Or _
               parsed(j) Like "*DR*" Or _
               parsed(j) Like "*DRQ*" Then
                Registri(i) = parsed(j)
                Exit For
            End If
        Next
    Next
    'odstrani vse duplikate
    registriUNI = RemoveDupesDict(Registri)
    
    
    Sheets("PDB").Cells(1, "A").Value = "[NodeName : " & SCADAname
    Sheets("PDB").Cells(1, "B").Value = "Database : " & SCADAname
    Sheets("PDB").Cells(1, "C").Value = "File Name : " & PLCName & "_" & SCADAname
    Sheets("PDB").Cells(1, "D").Value = "Date : " & Date
    Sheets("PDB").Cells(1, "E").Value = "Time : " & Time & "]"
    
    Sheets("PDB").Cells(2, "A").Value = "DR"
    Sheets("PDB").Cells(2, "B").Value = "AR"
    
    ' ANALOG REG
    ARglava1 = Array("[BLOCK TYPE", "TAG", "DESCRIPTION", "I/O DEVICE", "H/W OPTIONS", "I/O ADDRESS TYPE", "I/O ADDRESS", "SIGNAL CONDITIONING", "LOW EGU LIMIT", "HIGH EGU LIMIT", "EGU TAG", "OUTPUT ENABLE", "EVENT MESSAGES", "ALARM AREA(S)", "SECURITY AREA 1", "SECURITY AREA 2", "SECURITY AREA 3", "ALARM AREA 1", "ALARM AREA 2", "ALARM AREA 3", "ALARM AREA 4", "ALARM AREA 5", "ALARM AREA 6", "ALARM AREA 7", _
    "ALARM AREA 8", "ALARM AREA 9", "ALARM AREA 10", "ALARM AREA 11", "ALARM AREA 12", "ALARM AREA 13", "ALARM AREA 14", "ALARM AREA 15", "USER FIELD 1", "USER FIELD 2", "ESIG TYPE", "ESIG ALLOW CONT USE", "ESIG XMPT ALARM ACK", "ESIG UNSIGNED WRITES", "ESIG COMMENT REQUIRED", "PDR Update Rate", "PDR Access Time", "PDR Deadband", "PDR Latch", "PDR Disable Output", "PDR Array Length", "Hist Description", _
    "Hist Collect", "Hist Interval", "Hist Offset", "Hist Time Res", "Hist Compress", "Hist Deadband", "Hist Comp Type", "Hist Comp Time", "Scale Enabled", "Scale Clamping", "Scale Use EGU", "Scale Raw Low", "Scale Raw High", "Scale Low", "Scale High]")
    
    ARglava2 = Array("!A_NAME", "A_TAG", "A_DESC", "A_IODV", "A_IOHT", "A_NUMS", "A_IOAD", "A_IOSC", "A_ELO", "A_EHI", "A_EGUDESC", "A_OUT", "A_EVENT", "A_ADI", "A_SA1", "A_SA2", "A_SA3", "A_AREA1", "A_AREA2", "A_AREA3", "A_AREA4", "A_AREA5", "A_AREA6", "A_AREA7", "A_AREA8", "A_AREA9", "A_AREA10", "A_AREA11", "A_AREA12", "A_AREA13", "A_AREA14", "A_AREA15", "A_ALMEXT1", "A_ALMEXT2", "A_ESIGTYPE", "A_ESIGCONT", _
    "A_ESIGACK", "A_ESIGTRAP", "A_ESIGREQ_COMMENT", "A_PDR_UPDATERATE", "A_PDR_ACCESSTIME", "A_PDR_DEADBAND", "A_PDR_LATCHDATA", "A_PDR_DISABLEOUT", "A_PDR_ARRAYLENGTH", "A_HIST_DESC", "A_HIST_COLLECT", "A_HIST_INTERVAL", "A_HIST_OFFSET", "A_HIST_TIMERES", "A_HIST_COMPRESS", "A_HIST_DEADBAND", "A_HIST_COMPTYPE", "A_HIST_COMPTIME", "A_SCALE_ENABLED", "A_SCALE_CLAMP", "A_SCALE_USEEGU", "A_SCALE_RAWLOW", _
    "A_SCALE_RAWHIGH", "A_SCALE_LOW", "A_SCALE_HIGH!")
    
    
    For el = LBound(ARglava1) To UBound(ARglava1)
        Sheets("PDB").Cells(4, el + 1).Value = ARglava1(el)
        Sheets("PDB").Cells(5, el + 1).Value = ARglava2(el)
    Next
    nextFreeSpace = getNextRowNumber("PDB", "A", "!A_NAME") + 1
    For numberOfRegLines = LBound(registriUNI) To UBound(registriUNI)
        If registriUNI(numberOfRegLines) Like "*AR*" Then
            Sheets("PDB").Cells(nextFreeSpace, "A").Value = "AR"
            Sheets("PDB").Cells(nextFreeSpace, "B").Value = registriUNI(numberOfRegLines)
            Sheets("PDB").Cells(nextFreeSpace, "C").Value = PLCName & ": Analog register " & GetRegisterFromPDBname(registriUNI(numberOfRegLines))
            Sheets("PDB").Cells(nextFreeSpace, "D").Value = "GE9"
            If registriUNI(numberOfRegLines) Like "*DINT*" Or registriUNI(numberOfRegLines) Like "*REAL*" Then
                Sheets("PDB").Cells(nextFreeSpace, "E").Value = "ULong"
            Else
                Sheets("PDB").Cells(nextFreeSpace, "E").Value = ""
            End If
            Sheets("PDB").Cells(nextFreeSpace, "F").Value = "DECIMAL"
            Sheets("PDB").Cells(nextFreeSpace, "G").Value = PLCName & ":" & GetStartAddress(GetRegisterFromPDBname(registriUNI(numberOfRegLines)))
            If registriUNI(numberOfRegLines) Like "*DINT*" Or registriUNI(numberOfRegLines) Like "*REAL*" Then
                Sheets("PDB").Cells(nextFreeSpace, "H").Value = "None"
            Else
                Sheets("PDB").Cells(nextFreeSpace, "H").Value = "Lin"
            End If
            Sheets("PDB").Cells(nextFreeSpace, "I").Value = "-327,68"
            Sheets("PDB").Cells(nextFreeSpace, "I").Interior.ColorIndex = 3
            Sheets("PDB").Cells(nextFreeSpace, "J").Value = "327,67"
            Sheets("PDB").Cells(nextFreeSpace, "J").Interior.ColorIndex = 3
            Sheets("PDB").Cells(nextFreeSpace, "L").Value = "YES"
            Sheets("PDB").Cells(nextFreeSpace, "M").Value = "DISABLE"
            Sheets("PDB").Cells(nextFreeSpace, "N").Value = "NONE"
            Sheets("PDB").Cells(nextFreeSpace, "O").Value = "NONE"
            Sheets("PDB").Cells(nextFreeSpace, "P").Value = "NONE"
            Sheets("PDB").Cells(nextFreeSpace, "Q").Value = "NONE"
            Sheets("PDB").Cells(nextFreeSpace, "R").Value = "ALL"
            Sheets("PDB").Cells(nextFreeSpace, "AI").Value = "NONE"
            Sheets("PDB").Cells(nextFreeSpace, "AJ").Value = "YES"
            Sheets("PDB").Cells(nextFreeSpace, "AK").Value = "NO"
            Sheets("PDB").Cells(nextFreeSpace, "AL").Value = "REJECT"
            Sheets("PDB").Cells(nextFreeSpace, "AM").Value = "NO"
            Sheets("PDB").Cells(nextFreeSpace, "AN").Value = "1.000"
            Sheets("PDB").Cells(nextFreeSpace, "AO").Value = "300.000"
            Sheets("PDB").Cells(nextFreeSpace, "AP").Value = "0"
            Sheets("PDB").Cells(nextFreeSpace, "AQ").Value = "NO"
            Sheets("PDB").Cells(nextFreeSpace, "AR").Value = "NO"
            Sheets("PDB").Cells(nextFreeSpace, "AS").Value = "0"
            Sheets("PDB").Cells(nextFreeSpace, "AU").Value = "NO"
            Sheets("PDB").Cells(nextFreeSpace, "AV").Value = "5.000,00"
            Sheets("PDB").Cells(nextFreeSpace, "AW").Value = "0"
            Sheets("PDB").Cells(nextFreeSpace, "AX").Value = "Milliseconds"
            Sheets("PDB").Cells(nextFreeSpace, "AY").Value = "DISABLE"
            Sheets("PDB").Cells(nextFreeSpace, "AZ").Value = "0"
            Sheets("PDB").Cells(nextFreeSpace, "BA").Value = "Absolute"
            Sheets("PDB").Cells(nextFreeSpace, "BB").Value = "0"
            Sheets("PDB").Cells(nextFreeSpace, "BC").Value = "NO"
            Sheets("PDB").Cells(nextFreeSpace, "BD").Value = "NO"
            Sheets("PDB").Cells(nextFreeSpace, "BE").Value = "YES"
            Sheets("PDB").Cells(nextFreeSpace, "BF").Value = "0"
            Sheets("PDB").Cells(nextFreeSpace, "BG").Value = "65.535,00"
            Sheets("PDB").Cells(nextFreeSpace, "BH").Value = "-327,68"
            Sheets("PDB").Cells(nextFreeSpace, "BI").Value = "-327,68"
            nextFreeSpace = nextFreeSpace + 1
        End If
    Next

    ' DIGITAL REG
    DRglava1 = Array("[BLOCK TYPE", "TAG", "DESCRIPTION", "I/O DEVICE", "H/W OPTIONS", "I/O ADDRESS TYPE", "I/O ADDRESS", "ENABLE OUTPUT", "INVERT OUTPUT", "OPEN TAG", "CLOSE TAG", "EVENT MESSAGES", "ALARM AREA(S)", "SECURITY AREA 1", "SECURITY AREA 2", "SECURITY AREA 3", "ALARM AREA 1", "ALARM AREA 2", "ALARM AREA 3", "ALARM AREA 4", "ALARM AREA 5", "ALARM AREA 6", "ALARM AREA 7", "ALARM AREA 8", "ALARM AREA 9", "ALARM AREA 10", "ALARM AREA 11", "ALARM AREA 12", "ALARM AREA 13", "ALARM AREA 14", "ALARM AREA 15", "USER FIELD 1", "USER FIELD 2", "ESIG TYPE", "ESIG ALLOW CONT USE", "ESIG XMPT ALARM ACK", "ESIG UNSIGNED WRITES", "ESIG COMMENT REQUIRED", "PDR Update Rate", "PDR Access Time", "PDR Deadband", "PDR Latch", "PDR Disable Output", "PDR Array Length", "Hist Description", "Hist Collect", "Hist Interval", "Hist Offset", "Hist Time Res", "Hist Compress", "Hist Deadband", "Hist Comp Type", "Hist Comp Time]")
    DRglava2 = Array("!A_NAME", "A_TAG", "A_DESC", "A_IODV", "A_IOHT", "A_NUMS", "A_IOAD", "A_OUT", "A_INV", "A_OPENDESC", "A_CLOSEDESC", "A_EVENT", "A_ADI", "A_SA1", "A_SA2", "A_SA3", "A_AREA1", "A_AREA2", "A_AREA3", "A_AREA4", "A_AREA5", "A_AREA6", "A_AREA7", "A_AREA8", "A_AREA9", "A_AREA10", "A_AREA11", "A_AREA12", "A_AREA13", "A_AREA14", "A_AREA15", "A_ALMEXT1", "A_ALMEXT2", "A_ESIGTYPE", "A_ESIGCONT", "A_ESIGACK", "A_ESIGTRAP", "A_ESIGREQ_COMMENT", "A_PDR_UPDATERATE", "A_PDR_ACCESSTIME", "A_PDR_DEADBAND", "A_PDR_LATCHDATA", "A_PDR_DISABLEOUT", "A_PDR_ARRAYLENGTH", "A_HIST_DESC", "A_HIST_COLLECT", "A_HIST_INTERVAL", "A_HIST_OFFSET", "A_HIST_TIMERES", "A_HIST_COMPRESS", "A_HIST_DEADBAND", "A_HIST_COMPTYPE", "A_HIST_COMPTIME!")
      
    For el = LBound(DRglava1) To UBound(DRglava1)
        Sheets("PDB").Cells(nextFreeSpace + 1, el + 1).Value = DRglava1(el)
        Sheets("PDB").Cells(nextFreeSpace + 2, el + 1).Value = DRglava2(el)
    Next
    
    nextFreeSpace = nextFreeSpace + 3
    For numberOfRegLines = LBound(registriUNI) To UBound(registriUNI)
        If registriUNI(numberOfRegLines) Like "*DR*" Then
            Sheets("PDB").Cells(nextFreeSpace, "A").Value = "DR"
            Sheets("PDB").Cells(nextFreeSpace, "B").Value = registriUNI(numberOfRegLines)
            Sheets("PDB").Cells(nextFreeSpace, "C").Value = PLCName & ": Digital register " & GetRegisterFromPDBname(registriUNI(numberOfRegLines))
            Sheets("PDB").Cells(nextFreeSpace, "D").Value = "GE9"
            Sheets("PDB").Cells(nextFreeSpace, "F").Value = "DECIMAL"
            Sheets("PDB").Cells(nextFreeSpace, "G").Value = PLCName & ":" & GetStartAddress(GetRegisterFromPDBname(registriUNI(numberOfRegLines)))
            Sheets("PDB").Cells(nextFreeSpace, "H").Value = "YES"
            Sheets("PDB").Cells(nextFreeSpace, "I").Value = "NO"
            Sheets("PDB").Cells(nextFreeSpace, "J").Value = "OPEN"
            Sheets("PDB").Cells(nextFreeSpace, "K").Value = "CLOSE"
            Sheets("PDB").Cells(nextFreeSpace, "L").Value = "DISABLE"
            Sheets("PDB").Cells(nextFreeSpace, "M").Value = "NONE"
            Sheets("PDB").Cells(nextFreeSpace, "N").Value = "NONE"
            Sheets("PDB").Cells(nextFreeSpace, "O").Value = "NONE"
            Sheets("PDB").Cells(nextFreeSpace, "P").Value = "NONE"
            Sheets("PDB").Cells(nextFreeSpace, "Q").Value = "ALL"
            Sheets("PDB").Cells(nextFreeSpace, "AH").Value = "NONE"
            Sheets("PDB").Cells(nextFreeSpace, "AI").Value = "YES"
            Sheets("PDB").Cells(nextFreeSpace, "AJ").Value = "NO"
            Sheets("PDB").Cells(nextFreeSpace, "AK").Value = "REJECT"
            Sheets("PDB").Cells(nextFreeSpace, "AL").Value = "NO"
            Sheets("PDB").Cells(nextFreeSpace, "AM").Value = "1.000"
            Sheets("PDB").Cells(nextFreeSpace, "AN").Value = "300.000"
            Sheets("PDB").Cells(nextFreeSpace, "AO").Value = "0"
            Sheets("PDB").Cells(nextFreeSpace, "AP").Value = "NO"
            Sheets("PDB").Cells(nextFreeSpace, "AQ").Value = "NO"
            Sheets("PDB").Cells(nextFreeSpace, "AR").Value = "0"
            Sheets("PDB").Cells(nextFreeSpace, "AT").Value = "NO"
            Sheets("PDB").Cells(nextFreeSpace, "AU").Value = "5.000,00"
            Sheets("PDB").Cells(nextFreeSpace, "AV").Value = "0"
            Sheets("PDB").Cells(nextFreeSpace, "AW").Value = "Milliseconds"
            Sheets("PDB").Cells(nextFreeSpace, "AX").Value = "DISABLE"
            Sheets("PDB").Cells(nextFreeSpace, "AY").Value = "0"
            Sheets("PDB").Cells(nextFreeSpace, "AZ").Value = "Absolute"
            Sheets("PDB").Cells(nextFreeSpace, "BA").Value = "0"
            nextFreeSpace = nextFreeSpace + 1
        End If
        
    Next
    Sheets("PDB").Cells(nextFreeSpace + 1, "A").Value = EndLine
    msg = MsgBox("PDB import pripravljen za uvoz, kreirani so samo registri iz GE9 zavihka!", vbInformation, "Konec!")
End Sub

