Attribute VB_Name = "KreirajGE9"
Sub GE9_Create()
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
Dim Registri() As String
Dim registriUNI() As Variant
Dim parsed() As String
Dim parsedNext() As String

    'we delete and then create all the sheets
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("GE9").Delete
    Sheets.Add.Name = "GE9"
    PLCName = Sheets("IOT").Cells(1, "I").Value
    
    ' oèisti TGD substitutione tako da dobiš samo registre
    numberOfElements = Sheets("TGD").Cells(1, "A").Value
    ReDim Registri(numberOfElements)
    For i = LBound(Registri) To UBound(Registri)
        parsed = Split(Sheets("TGD").Cells(i, "B").Value, ".")
        For j = LBound(parsed) To UBound(parsed)
            If parsed(j) Like "*AR*" Or _
               parsed(j) Like "*DR*" Or _
               parsed(j) Like "*DRQ*" Then
                parsedNext = Split(parsed(j), "_")
                For k = LBound(parsedNext) To UBound(parsedNext)
                    If parsedNext(k) Like "*AR*" Or _
                       parsedNext(k) Like "*DR*" Or _
                       parsedNext(k) Like "*DRQ*" Then
                        Registri(i) = parsedNext(k)
                        Exit For
                    End If
                Next
            End If
        Next
    Next
    'odstrani vse duplikate
    registriUNI = RemoveDupesDict(Registri)
    
    'naredimo fiksne vrstice katere nimajo veze z INPUT podatki 'Tuesday December 27 2022, 09:08 AM
    Sheets("GE9").Cells(1, "A").Value = "[GE9 I/O Driver Configuration Report, " & Format(Date, "dddd, mmm d yyyy") & ", " & Format(Time, "hh:mm:ss AM/PM") & "]"
    Sheets("GE9").Cells(3, "A").Value = "!Name"
    Sheets("GE9").Cells(3, "B").Value = "Description"
    Sheets("GE9").Cells(3, "C").Value = "Enabled"
    
    Sheets("GE9").Cells(4, "A").Value = "Kanal1"
    Sheets("GE9").Cells(4, "B").Value = ""
    Sheets("GE9").Cells(4, "C").Value = "1"
    
    Sheets("GE9").Cells(6, "A").Value = "@Channel"
    Sheets("GE9").Cells(6, "B").Value = "Name"
    Sheets("GE9").Cells(6, "C").Value = "Description"
    Sheets("GE9").Cells(6, "D").Value = "Enabled"
    Sheets("GE9").Cells(6, "E").Value = "PrimaryIpAddress"
    Sheets("GE9").Cells(6, "F").Value = "PrimaryReplyTimeout"
    Sheets("GE9").Cells(6, "G").Value = "PrimaryRetries"
    Sheets("GE9").Cells(6, "H").Value = "PrimaryDelay"
    Sheets("GE9").Cells(6, "I").Value = "BackupIpAddress"
    Sheets("GE9").Cells(6, "J").Value = "BackupReplyTimeout"
    Sheets("GE9").Cells(6, "K").Value = "BackupRetries"
    Sheets("GE9").Cells(6, "L").Value = "BackupDelay"
    Sheets("GE9").Cells(6, "M").Value = "TcpOrUdp"
    Sheets("GE9").Cells(6, "N").Value = "Password"
    Sheets("GE9").Cells(6, "O").Value = "PrivilegeLevel"

    Sheets("GE9").Cells(7, "A").Value = Sheets("GE9").Cells(4, "A").Value
    Sheets("GE9").Cells(7, "B").Value = PLCName
    Sheets("GE9").Cells(7, "C").Value = "Krmilnik " & PLCName
    Sheets("GE9").Cells(7, "D").Value = "1"
    Sheets("GE9").Cells(7, "E").Value = "Vstavi IP!!"
    Sheets("GE9").Cells(7, "E").Interior.ColorIndex = 3
    Sheets("GE9").Cells(7, "F").Value = "1"
    Sheets("GE9").Cells(7, "G").Value = "3"
    Sheets("GE9").Cells(7, "H").Value = "30"
    Sheets("GE9").Cells(7, "I").Value = ""
    Sheets("GE9").Cells(7, "J").Value = "1"
    Sheets("GE9").Cells(7, "K").Value = "3"
    Sheets("GE9").Cells(7, "L").Value = "30"
    Sheets("GE9").Cells(7, "M").Value = "1"
    Sheets("GE9").Cells(7, "N").Value = ""
    Sheets("GE9").Cells(7, "O").Value = "0"

    Sheets("GE9").Cells(9, "A").Value = "#Device"
    Sheets("GE9").Cells(9, "B").Value = "Name"
    Sheets("GE9").Cells(9, "C").Value = "Description"
    Sheets("GE9").Cells(9, "D").Value = "StartAddress"
    Sheets("GE9").Cells(9, "E").Value = "Length"
    Sheets("GE9").Cells(9, "F").Value = "PrimaryPollTime"
    Sheets("GE9").Cells(9, "G").Value = "SecondaryPollTime"
    Sheets("GE9").Cells(9, "H").Value = "Phase"
    Sheets("GE9").Cells(9, "I").Value = "AccessTime"
    Sheets("GE9").Cells(9, "J").Value = "DeadBand"
    Sheets("GE9").Cells(9, "K").Value = "Enabled"
    Sheets("GE9").Cells(9, "L").Value = "LatchData"
    Sheets("GE9").Cells(9, "M").Value = "OutputDisabled"
    Sheets("GE9").Cells(9, "N").Value = "BlockWritesEnabled"
    Sheets("GE9").Cells(9, "O").Value = "DataType"
    
    nextFreeSpace = getNextRowNumber("GE9", "A", "#Device") + 1
    
    For j = LBound(registriUNI) To UBound(registriUNI)
        If Not registriUNI(j) Like "" Then
            
            Sheets("GE9").Cells(nextFreeSpace, "A").Value = PLCName
            Sheets("GE9").Cells(nextFreeSpace, "B").Value = PLCName & "_" & registriUNI(j)
            
            If registriUNI(j) Like "*AR*" Then
                Sheets("GE9").Cells(nextFreeSpace, "C").Value = "Analogni register " & registriUNI(j)
            ElseIf registriUNI(j) Like "*DR*" Then
                Sheets("GE9").Cells(nextFreeSpace, "C").Value = "Digitalni register " & registriUNI(j)
            End If
        
            Sheets("GE9").Cells(nextFreeSpace, "D").Value = GetStartAddress(registriUNI(j))
      
            If registriUNI(j) Like "*AR*" Then
                Sheets("GE9").Cells(nextFreeSpace, "E").Value = "300"
            ElseIf registriUNI(j) Like "*DR*" Then
                Sheets("GE9").Cells(nextFreeSpace, "E").Value = "1000"
            End If
            
            Sheets("GE9").Cells(nextFreeSpace, "F").Value = "1"
            Sheets("GE9").Cells(nextFreeSpace, "G").Value = ""
            Sheets("GE9").Cells(nextFreeSpace, "H").Value = "0"
            Sheets("GE9").Cells(nextFreeSpace, "I").Value = "05:00"
            Sheets("GE9").Cells(nextFreeSpace, "J").Value = "1"
            Sheets("GE9").Cells(nextFreeSpace, "K").Value = "1"
            Sheets("GE9").Cells(nextFreeSpace, "L").Value = "0"
            Sheets("GE9").Cells(nextFreeSpace, "M").Value = "0"
            Sheets("GE9").Cells(nextFreeSpace, "N").Value = "0"
            If registriUNI(j) Like "*AR*" Then
                Sheets("GE9").Cells(nextFreeSpace, "O").Value = "1"
            ElseIf registriUNI(j) Like "*DR*" Then
                Sheets("GE9").Cells(nextFreeSpace, "O").Value = "4"
            End If
            nextFreeSpace = nextFreeSpace + 1
        End If
    Next
    msg = MsgBox("Konfiguracija za gonilnik GE9 je pripravljena za kopiranje v GE9 power tool. Good Luck!", vbInformation, "Konec!")
End Sub
