Attribute VB_Name = "Kreiraj_PLC_logiko_MEJE"
Sub MEJE_Create()

Dim LC_Value As Double
Dim UC_Value As Double
Dim TagNameLC As String
Dim TagNameUC As String
Dim TagNameW As String
Dim TagNameK As String
Dim RowCounter As Integer
Dim AIrowNumber As Integer
Dim endLoop As Boolean
endLoop = False

'zbrišemo in ustvarimo Sheet
Application.DisplayAlerts = False
On Error Resume Next
Sheets("MEJE_PLC").Delete
Sheets.Add.Name = "MEJE_PLC"

RowCounter = 1
AIrowNumber = 1

    Do Until endLoop
        'End Loop
        If Sheets("IOT").Cells(RowCounter, "A").Value Like "" Then
            endLoop = True
            MsgBox "Kreiranje kode MEJE konèano, kopiraj vsebino zavihka 'MEJE_PLC' v PLC programski blok MEJE"
            Exit Sub
        End If
        'Najšo sm AI!!
        If Sheets("IOT").Cells(RowCounter, "A").Value Like "*%AI*" Then
            TagNameLC = Sheets("IOT").Cells(RowCounter, "B").Value & "_VA_" & Sheets("IOT").Cells(RowCounter, "C").Value & "_LC"
            TagNameUC = Sheets("IOT").Cells(RowCounter, "B").Value & "_VA_" & Sheets("IOT").Cells(RowCounter, "C").Value & "_UC"
            TagNameW = Sheets("IOT").Cells(RowCounter, "B").Value & "_VA_" & Sheets("IOT").Cells(RowCounter, "C").Value & "_WEIGHT"
            TagNameK = Sheets("IOT").Cells(RowCounter, "B").Value & "_VA_" & Sheets("IOT").Cells(RowCounter, "C").Value & "_KOR"
            If Sheets("IOT").Cells(RowCounter, "B").Value Like "REZ" Then
                LC_Value = 0
                UC_Value = 0
            Else
                LC_Value = Sheets("IOT").Cells(RowCounter, "E").Value
                Select Case LC_Value
                    Case Is > -327
                        LC_Value = 100 * LC_Value
                    Case -327 To -3276
                        LC_Value = 10 * LC_Value
                    Case Else
                        LC_Value = 1 * LC_Value
                End Select
                
                UC_Value = Sheets("IOT").Cells(RowCounter, "F").Value
                Select Case UC_Value
                    Case Is < 327
                        UC_Value = 100 * UC_Value
                    Case 327 To 3276
                        UC_Value = 10 * UC_Value
                    Case Else
                        UC_Value = 1 * UC_Value
                End Select
            End If
            
            ' ustvarjanje blokov v Sheet-u
            Sheets("MEJE_PLC").Cells(AIrowNumber, "A").Value = "COMMENT /* " & Sheets("IOT").Cells(RowCounter, "D").Value & _
            " */; END_RUNG;" & "H_WIRE;"
            Sheets("MEJE_PLC").Cells(AIrowNumber, "B").Value = "MOVE_INT 1 " & LC_Value & " " & TagNameLC & ";"
            Sheets("MEJE_PLC").Cells(AIrowNumber, "C").Value = "H_WIRE;"
            Sheets("MEJE_PLC").Cells(AIrowNumber, "D").Value = "H_WIRE;"
            Sheets("MEJE_PLC").Cells(AIrowNumber, "E").Value = "MOVE_INT 1 " & UC_Value & " " & TagNameUC & ";"
            Sheets("MEJE_PLC").Cells(AIrowNumber, "F").Value = "H_WIRE;"
            Sheets("MEJE_PLC").Cells(AIrowNumber, "G").Value = "H_WIRE;"
            Sheets("MEJE_PLC").Cells(AIrowNumber, "H").Value = "MOVE_INT 1 " & "10" & " " & TagNameW & ";"
            Sheets("MEJE_PLC").Cells(AIrowNumber, "I").Value = "H_WIRE;"
            Sheets("MEJE_PLC").Cells(AIrowNumber, "J").Value = "H_WIRE;"
            Sheets("MEJE_PLC").Cells(AIrowNumber, "K").Value = "MOVE_INT 1 " & "0" & " " & TagNameK & ";"
            Sheets("MEJE_PLC").Cells(AIrowNumber, "L").Value = "END_RUNG;"
            AIrowNumber = AIrowNumber + 1
            
        End If
     
        'Next
        RowCounter = RowCounter + 1
        
    Loop
End Sub
