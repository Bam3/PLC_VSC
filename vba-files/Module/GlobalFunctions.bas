Attribute VB_Name = "GlobalFunctions"
Public Function getNextRowNumber(inWhatSheetName As String, inWhatColumn As String, findWhat As String) As Long
    Sheets(inWhatSheetName).Select
    Columns(inWhatColumn & ":" & inWhatColumn).Select
    Selection.Find(findWhat).Select
    getNextRowNumber = Selection.Row
End Function
Public Function createCodeForPLC(ByVal systemName As String, _
                                 ByVal sensorName As String, _
                                 ByVal sensorDescription As String, _
                                 ByVal sensorType As String, _
                                 ByVal sheetName As String, _
                                 ByVal address As String, _
                                 ByVal PLCName As String)
    

    Sheets(sheetName).Select
    Range("A1").End(xlDown).Offset(0, 0).Select
    'prepre�imo pisanje izven worksheet-a
    If ActiveCell.Row >= 1048576 Then
        Sheets(sheetName).Cells(1, "A").Select
    End If
    'first line
    ActiveCell.Offset(0, 0).Value = "COMMENT /*" & sensorDescription & "*/;"
    ActiveCell.Offset(0, 1).Value = "END_RUNG;"
    
    '---------------------------CREATE CTRL_DQ---------------------------
    If sensorType Like "D" Then
        'third line
        ActiveCell.Offset(1, 0).Value = "NOCON #ALW_ON,G,;"
        ActiveCell.Offset(1, 1).Value = "CTRL_DQ " & sensorName & "_D" & ",L " _
                                        & systemName & "_VD_" & sensorName & "_OFF,G, " _
                                        & systemName & "_VD_" & sensorName & "_AU,G, " _
                                        & systemName & "_VD_" & sensorName & "_MN,G, " _
                                        & systemName & "_VD_" & sensorName & "_SR,G, " _
                                        & systemName & "_VD_" & sensorName & "_BL,G, " _
                                        & systemName & "_VD_" & sensorName & "_INV,G, " _
                                        & systemName & "_VD_VKLOP_SCADA,G, " _
                                        & systemName & "_VA_VODENJE,G, " _
                                        & systemName & "_VA_" & sensorName & "_RZ,G, " _
                                        & systemName & "_DO_" & sensorName & ",G, " _
                                        & systemName & "_VA_" & sensorName & "_S,G, ;"
        ActiveCell.Offset(1, 2).Value = "END_RUNG;"
    End If
    '---------------------------CREATE CTRL_DQ, CTRL_EGU2AQ, CHK_ACT---------------------------
    If sensorType Like "DAC" Then
        ActiveCell.Offset(1, 0).Value = "NOCON #ALW_ON,G,;"
        ActiveCell.Offset(1, 1).Value = "CTRL_DQ " & sensorName & "_D" & ",L " _
                                        & systemName & "_VD_" & sensorName & "_OFF,G, " _
                                        & systemName & "_VD_" & sensorName & "_AU,G, " _
                                        & systemName & "_VD_" & sensorName & "_MN,G, " _
                                        & systemName & "_VD_" & sensorName & "_SR,G, " _
                                        & systemName & "_VD_" & sensorName & "_BL,G, " _
                                        & systemName & "_VD_" & sensorName & "_INV,G, " _
                                        & systemName & "_VD_VKLOP_SCADA,G, " _
                                        & systemName & "_VA_VODENJE,G, " _
                                        & systemName & "_VA_" & sensorName & "_RZ,G, ** " _
                                        & systemName & "_VA_" & sensorName & "_S,G, ;"
        ActiveCell.Offset(1, 2).Value = "H_WIRE;"
        ActiveCell.Offset(1, 3).Value = "H_WIRE;"
        ActiveCell.Offset(1, 4).Value = "CTRL_EGU2AQ " & sensorName & "_A" & ",L " _
                                        & systemName & "_VA_" & sensorName & "_OFF,G, " _
                                        & systemName & "_VA_" & sensorName & "_AU,G, " _
                                        & systemName & "_VA_" & sensorName & "_MN,G, " _
                                        & systemName & "_VA_" & sensorName & "_SR,G, " _
                                        & systemName & "_VA_VODENJE,G, " _
                                        & systemName & "_VA_" & sensorName & "_RZ,G, " _
                                        & 0 & ",L " _
                                        & systemName & "_VD_" & sensorName & "_BL,G, " _
                                        & systemName & "_VD_VKLOP_SCADA,G, " _
                                        & systemName & "_AO_" & sensorName & ",G, " _
                                        & systemName & "_VA_" & sensorName & ",G, ** ;"
        ActiveCell.Offset(1, 5).Value = "H_WIRE;"
        ActiveCell.Offset(1, 6).Value = "H_WIRE;"
        ActiveCell.Offset(1, 7).Value = "CHK_ACT " & sensorName & "_C" & ",L " _
                                        & sensorName & "_D.Q,L " _
                                        & systemName & "_VD_XS_" & sensorName & ",G, " _
                                        & systemName & "_VD_XA_" & sensorName & ",G, " _
                                        & "#ALW_ON,G, " _
                                        & systemName & "_VD_KVIT_SCADA,G, " _
                                        & 120 & ",L " _
                                        & 5 & ",L " _
                                        & systemName & "_VD_" & sensorName & "_OBRHD_R,G, " _
                                        & systemName & "_VA_" & sensorName & "_OBRHD,G, " _
                                        & systemName & "_VD_" & sensorName & "_ST_VKL_R,G, " _
                                        & systemName & "_VA_" & sensorName & "_RZ,G, " _
                                        & systemName & "_DO_" & sensorName & ",G, " _
                                        & systemName & "_VD_" & sensorName & "_E_DEL,G, " _
                                        & systemName & "_VD_" & sensorName & "_E_JER,G, " _
                                        & systemName & "_VD_" & sensorName & "_E_FP,G, " _
                                        & systemName & "_VA_" & sensorName & "_ST_VKL,G, ;"
        ActiveCell.Offset(1, 8).Value = "END_RUNG;"
    End If
    '---------------------------CREATE CTRL_EGU2AQ---------------------------
    If sensorType Like "A" Then
        ActiveCell.Offset(1, 0).Value = "NOCON #ALW_ON,G,;"
        ActiveCell.Offset(1, 1).Value = "CTRL_EGU2AQ " & sensorName & "_A" & ",L " _
                                        & systemName & "_VA_" & sensorName & "_OFF,G, " _
                                        & systemName & "_VA_" & sensorName & "_AU,G, " _
                                        & systemName & "_VA_" & sensorName & "_MN,G, " _
                                        & systemName & "_VA_" & sensorName & "_SR,G, " _
                                        & systemName & "_VA_VODENJE,G, " _
                                        & systemName & "_VA_" & sensorName & "_RZ,G, " _
                                        & 0 & ",L " _
                                        & systemName & "_VD_" & sensorName & "_BL,G, " _
                                        & systemName & "_VD_VKLOP_SCADA,G, " _
                                        & systemName & "_AO_" & sensorName & ",G, " _
                                        & systemName & "_VA_" & sensorName & ",G, " _
                                        & systemName & "_VA_" & sensorName & "_S,G, ;"
        ActiveCell.Offset(1, 2).Value = "END_RUNG;"
    End If
    '---------------------------CREATE CTRL_DQ in CHK_ACT---------------------------
    If sensorType Like "DC" Then
        ActiveCell.Offset(1, 0).Value = "NOCON #ALW_ON,G,;"
        ActiveCell.Offset(1, 1).Value = "CTRL_DQ " & sensorName & "_D" & ",L " _
                                        & systemName & "_VD_" & sensorName & "_OFF,G, " _
                                        & systemName & "_VD_" & sensorName & "_AU,G, " _
                                        & systemName & "_VD_" & sensorName & "_MN,G, " _
                                        & systemName & "_VD_" & sensorName & "_SR,G, " _
                                        & systemName & "_VD_" & sensorName & "_BL,G, " _
                                        & systemName & "_VD_" & sensorName & "_INV,G, " _
                                        & systemName & "_VD_VKLOP_SCADA,G, " _
                                        & systemName & "_VA_VODENJE,G, " _
                                        & systemName & "_VA_" & sensorName & "_RZ,G, " _
                                        & systemName & "_DO_" & sensorName & ",G, " _
                                        & systemName & "_VA_" & sensorName & "_S,G, ;"
        ActiveCell.Offset(1, 2).Value = "H_WIRE;"
        ActiveCell.Offset(1, 3).Value = "H_WIRE;"
        ActiveCell.Offset(1, 4).Value = "CHK_ACT " & sensorName & "_C" & ",L " _
                                        & sensorName & "_D.Q,L " _
                                        & systemName & "_VD_XS_" & sensorName & ",G, " _
                                        & systemName & "_VD_XA_" & sensorName & ",G, " _
                                        & "#ALW_ON,G, " _
                                        & systemName & "_VD_KVIT_SCADA,G, " _
                                        & 120 & ",L " _
                                        & 5 & ",L " _
                                        & systemName & "_VD_" & sensorName & "_OBRHD_R,G, " _
                                        & systemName & "_VA_" & sensorName & "_OBRHD,G, " _
                                        & systemName & "_VD_" & sensorName & "_ST_VKL_R,G, " _
                                        & systemName & "_VA_" & sensorName & "_RZ,G, " _
                                        & systemName & "_DO_" & sensorName & ",G, " _
                                        & systemName & "_VD_" & sensorName & "_E_DEL,G, " _
                                        & systemName & "_VD_" & sensorName & "_E_JER,G, " _
                                        & systemName & "_VD_" & sensorName & "_E_FP,G, " _
                                        & systemName & "_VA_" & sensorName & "_ST_VKL,G, ;"
        ActiveCell.Offset(1, 5).Value = "END_RUNG;"
    End If
    '---------------------------CREATE I2VD---------------------------
    If sensorType Like "I2VD" Then
        ActiveCell.Offset(1, 0).Value = "NOCON #ALW_ON,G,;"
        ActiveCell.Offset(1, 1).Value = "I2VD " & address & ",L " _
                                        & systemName & "_DI_" & sensorName & ",G, " _
                                        & systemName & "_VD_" & sensorName & "_SB,G, " _
                                        & "#ALW_OFF,G, " _
                                        & systemName & "_VD_" & sensorName & "_SV,G, " _
                                        & systemName & "_VD_" & sensorName & ",G, ;"
        ActiveCell.Offset(1, 2).Value = "END_RUNG;"
    End If
    
    '---------------------------CREATE AI2EGU---------------------------
    If sensorType Like "AI2EGU" Then
    
        'get number of AI module for Error name variable
        errorName = CInt(Replace(address, "AI", ""))
        If (errorName Mod 16 = 0) Then
            errorName = Int(errorName / 16)
        Else
            errorName = Int(errorName / 16) + 1
        End If
    
        ActiveCell.Offset(1, 0).Value = "NOCON #ALW_ON,G,;"
        ActiveCell.Offset(1, 1).Value = "AI2EGU_PAC " & address & ",L " _
                                        & systemName & "_AI_" & sensorName & ",G, " _
                                        & systemName & "_VA_" & sensorName & "_LC,G, " _
                                        & systemName & "_VA_" & sensorName & "_UC,G, " _
                                        & systemName & "_VA_" & sensorName & "_WEIGHT,G, " _
                                        & systemName & "_VA_" & sensorName & "_KOR,G, " _
                                        & PLCName & "_T_AI_MODULE_" & Format(errorName, "00") & "_ERR,G, " _
                                        & systemName & "_VA_" & sensorName & ",G, " _
                                        & systemName & "_VD_" & sensorName & "_E_SENS,G, ;"
        ActiveCell.Offset(1, 2).Value = "END_RUNG;"
    End If
    
    '---------------------------CREATE EGUALM---------------------------
    If sensorType Like "EGUALM" Then
        If sensorName Like "AI*" Then
            sysAlmEnable = "** "
        ElseIf sensorName Like "*P*" Then
            sysAlmEnable = systemName & "_VD_AL_ENABLE_P,G, "
        Else
            sysAlmEnable = systemName & "_VD_AL_ENABLE_TH,G, "
        End If
        ActiveCell.Offset(1, 0).Value = "NOCON #ALW_ON,G,;"
        ActiveCell.Offset(1, 1).Value = "EGU_4AL_PAC " & address & ",L " _
                                        & sysAlmEnable _
                                        & systemName & "_VD_" & sensorName & "_A_EN,G, " _
                                        & systemName & "_VD_" & sensorName & "_KVIT,G, " _
                                        & systemName & "_VA_" & sensorName & ",G, " _
                                        & systemName & "_VA_" & sensorName & "_HIHI,G, " _
                                        & systemName & "_VA_" & sensorName & "_HI,G, " _
                                        & systemName & "_VA_" & sensorName & "_LO,G, " _
                                        & systemName & "_VA_" & sensorName & "_LOLO,G, " _
                                        & systemName & "_VA_" & sensorName & "_ZAK1,G, " _
                                        & systemName & "_VA_" & sensorName & "_ZAK2,G, " _
                                        & systemName & "_VD_" & sensorName & "_A_HIHI,G, " _
                                        & systemName & "_VD_" & sensorName & "_A_HI,G, " _
                                        & systemName & "_VD_" & sensorName & "_A_LO,G, " _
                                        & systemName & "_VD_" & sensorName & "_A_LOLO,G, ;"
        ActiveCell.Offset(1, 2).Value = "END_RUNG;"
    End If
    
    ActiveCell.Offset(2, 0).Value = " "
    ActiveCell.Offset(3, 0).Value = " "


End Function
Sub saveAsCSVslo(sheetName As String, fileName As String, directory As String)


Dim sFileName As String
Dim WB As Workbook

Application.DisplayAlerts = False
sFileName = fileName & ".csv"

    'Copy the contents of required sheet ready to paste into the new CSV
    Sheets(sheetName).UsedRange.Copy
    
    'Open a new XLS workbook, save it as the file name
    Set WB = Workbooks.Add
    With WB
        .Title = fileName
        .Sheets(1).Select
        ActiveSheet.Paste
        .SaveAs directory & sFileName, xlCSV, Local:=False
        .Close
    End With
    
Application.DisplayAlerts = True

heh = MsgBox("Datoteka " & sFileName & " se nahaja na lokaciji: " & directory, vbInformation, "Sporo�ilo!")

End Sub

Public Function decimalneVejiceEnote(ByVal enotaSenzorja As String, ByVal mejaSenzorja As String) As String

    If enotaSenzorja Like "�C" _
    Or enotaSenzorja Like "%" _
    Or enotaSenzorja Like "m3/h" Then
        decimalneVejiceEnote = Format(mejaSenzorja, "#0.0")
    ElseIf enotaSenzorja Like "bar" Then
        decimalneVejiceEnote = Format(mejaSenzorja, "#0.00")
    ElseIf enotaSenzorja Like "kg/h" Then
        decimalneVejiceEnote = Format(mejaSenzorja, "#0")
    Else 'default je 1 dec. mesto
        decimalneVejiceEnote = Format(mejaSenzorja, "#0.0")
    End If
    
    '�e je lo�ilna oznaka "." in ne vejica!
    If InStr(decimalneVejiceEnote, ".") Then
        decimalneVejiceEnote = Replace(decimalneVejiceEnote, ".", ",")
    End If

End Function
Function RemoveDupesDict(MyArray As Variant) As Variant
'DESCRIPTION: Removes duplicates from your array using the dictionary method.
'NOTES: (1.a) You must add a reference to the Microsoft Scripting Runtime library via
' the Tools > References menu.
' (1.b) This is necessary because I use Early Binding in this function.
' Early Binding greatly enhances the speed of the function.
' (2) The scripting dictionary will not work on the Mac OS.
'SOURCE: https://wellsr.com
'-----------------------------------------------------------------------
    Dim i As Long
    Set d = CreateObject("Scripting.Dictionary")
    'Dim d As Scripting.Dictionary
    'Set d = New Scripting.Dictionary
    With d
        For i = LBound(MyArray) To UBound(MyArray)
            If IsMissing(MyArray(i)) = False Then
                .Item(MyArray(i)) = 1
            End If
        Next
        RemoveDupesDict = .Keys
    End With
End Function

Function GetStartAddress(ByVal register As String) As String
Dim midAcc() As String
Dim numberOfRegister As Integer

If register Like "DRQ*" Then
    midAcc = Split(register, "DRQ")
    numberOfRegister = CInt(midAcc(1))
    GetStartAddress = "Q" & numberOfRegister
    Exit Function
End If

If register Like "DR*" Then
    midAcc = Split(register, "DR")
    numberOfRegister = CInt(midAcc(1))
    GetStartAddress = "M" & Format(numberOfRegister - 1 & "001", "")
End If

If register Like "AR*" Then
    midAcc = Split(register, "AR")
    numberOfRegister = CInt(midAcc(1))
    GetStartAddress = "R" & Format((numberOfRegister - 1) * 3 & "01", "")
End If

End Function


Function GetRegisterFromPDBname(ByVal register As String) As String
Dim midAcc() As String
Dim numberOfRegister As Integer
    
    midAcc = Split(register, "_")
    For Counter = LBound(midAcc) To UBound(midAcc)
        If midAcc(Counter) Like "*AR*" Then
            GetRegisterFromPDBname = midAcc(Counter)
            Exit Function
        End If
        If midAcc(Counter) Like "*DR*" Then
            GetRegisterFromPDBname = midAcc(Counter)
            Exit Function
        End If
        If midAcc(Counter) Like "*DRQ*" Then
            GetRegisterFromPDBname = midAcc(Counter)
            Exit Function
        End If
    Next
End Function

Public Sub RibbonLook()
    MsgBox "Ribbon", , "Ribbon"
End Sub
