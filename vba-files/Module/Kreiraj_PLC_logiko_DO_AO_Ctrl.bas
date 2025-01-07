Attribute VB_Name = "Kreiraj_PLC_logiko_DO_AO_Ctrl"
Sub CTRL_Create()
Dim codeTypes As Object
Set codeTypes = CreateObject("Scripting.Dictionary")
Dim numberOfElements As Long
Dim Description As String
Dim system As String
Dim tag As String
Dim i As Long
Dim msg As Boolean
Dim sheetName As String
Dim address As String
Dim PLCName As String

    'we set the parser values for codeType
    codeTypes.Add Key:="D", Item:="_DO_"
    codeTypes.Add Key:="DC", Item:="_DO_"
    codeTypes.Add Key:="DAC", Item:="_DO_"
    codeTypes.Add Key:="A", Item:="_AO_"
    
    
    'we delete and then create all the sheets
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("CTRL_PLC").Delete
    Sheets.Add.Name = "CTRL_PLC"
    PLCName = Sheets("IOT").Cells(4, "F").Value
    'get all the elements from "vhodnaTabela"
    numberOfElements = getNextRowNumber("IOT", "A", "") - 1
    
    'create CTRL
    For i = 2 To numberOfElements
        If Sheets("IOT").Cells(i, "A").Value Like "*Q*" And Not Sheets("IOT").Cells(i, "E").Value Like "" Then
            'check for CODE TYPE
            address = Replace(Sheets("IOT").Cells(i, "A").Value, "%", "")
            codeType = Sheets("IOT").Cells(i, "H").Value
            system = Sheets("IOT").Cells(i, "B").Value
            tag = Sheets("IOT").Cells(i, "C").Value
            Description = system & ": " & Sheets("IOT").Cells(i, "D").Value
            Call createCodeForPLC(system, tag, Description, codeType, "CTRL_PLC", address, PLCName)
        End If
    Next i
    msg = MsgBox("Koda CTRL je pripravljena za kopiranje v PAC Machine Edition", vbInformation, "Konec!")

End Sub






