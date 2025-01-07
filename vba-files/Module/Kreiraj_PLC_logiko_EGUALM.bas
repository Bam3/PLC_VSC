Attribute VB_Name = "Kreiraj_PLC_logiko_EGUALM"
Sub EGUALM_Create()
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


    'we delete and then create all the sheets
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("EGUALM_PLC").Delete
    Sheets.Add.Name = "EGUALM_PLC"

    'get all the elements from "vhodnaTabela"
    numberOfElements = getNextRowNumber("IOT", "A", "") - 1
    PLCName = Sheets("IOT").Cells(4, "F").Value
    'create CTRL
    For i = 2 To numberOfElements
        If Sheets("IOT").Cells(i, "A").Value Like "%AI*" Then
            address = Replace(Sheets("IOT").Cells(i, "A").Value, "%", "")
            system = Sheets("IOT").Cells(i, "B").Value
            tag = Sheets("IOT").Cells(i, "C").Value
            Description = system & ": " & Sheets("IOT").Cells(i, "D").Value
            Call createCodeForPLC(system, tag, Description, "EGUALM", "EGUALM_PLC", address, PLCName)
        End If
    Next i
    msg = MsgBox("Koda AL_AI je pripravljena za kopiranje v PAC Machine Edition", vbInformation, "Konec!")

End Sub


