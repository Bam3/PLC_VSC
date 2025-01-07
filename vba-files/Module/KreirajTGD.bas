Attribute VB_Name = "KreirajTGD"
Public Sub createTGD()
    UserFormMakeTGD.Show
End Sub
Public Function GetNumLoc(xValue As String) As Integer

For GetNumLoc = 1 To Len(xValue)
    If Mid(xValue, GetNumLoc, 1) Like "#" Then Exit Function
Next
GetNumLoc = 0

End Function

Sub exportPLC_to_TDG(ByVal nodeName, _
                     ByVal PLCName, _
                     ByVal registerRcount, _
                     ByVal registerMcount, _
                     ByVal registerQcount, _
                     Predpona As String, _
                     zacAO As String, _
                     koncAO As String, _
                     ByVal suffixEnable, _
                     ByVal Security, _
                     ByVal registerFormat, _
                     ByVal registerFormatDRQ, _
                     ByVal InputSheet)
'-------------------------------------------------

Dim rw As Range
Dim CSVname As String, temporary As String, Default As String
Dim temp() As String
Dim cnt As Integer, x As Integer, y As Integer, z As Integer
Dim registerFormatPrint 'As Integer
Dim registerFormatPrintDRQ As String
Dim obmocjeTEHNIK As String
Dim obmocjeTEHNOLOG As String

'doloèitev security area
obmocjeTEHNIK = "A"
obmocjeTEHNOLOG = "B"

'postavimo števec na 0
cnt = 0

'skrijemo formo
UserFormMakeTGD.Hide

'preverimo vnosna polja
If PLCName = vbNullString Or nodeName = vbNullString Then
    MsgBox "Prekinjeno!?"
    Exit Sub
End If

'we delete and then create the sheet
Application.DisplayAlerts = False
On Error Resume Next
Sheets("TGD").Delete
Sheets.Add.Name = "TGD"

For Each rw In Worksheets(InputSheet).Rows
    If Worksheets(InputSheet).Cells(rw.Row, 1).Value = "" Or Worksheets(InputSheet).Cells(rw.Row, 1).Value = "IOAddress" Then
        Exit For
    End If
    temporary = GetNumLoc(Worksheets(InputSheet).Cells(rw.Row, 15).Value)
    If Not temporary = 0 Then
        If temporary = 4 Then
            x = CInt(Right(Worksheets(InputSheet).Cells(rw.Row, 15).Value, 4))
        End If
        If temporary = 3 Then
            x = CInt(Right(Worksheets(InputSheet).Cells(rw.Row, 15).Value, 5))
        End If
    End If
            
'Dejanski naslov toèke
addressa = x
            
'R (do 1500)------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If InStr(Worksheets(InputSheet).Cells(rw.Row, 15).Value, "%R") > 0 And x < CInt(registerRcount) + 1 Then
        DataType = Worksheets(InputSheet).Cells(rw.Row, 2).Value
        cnt = cnt + 1
        y = x \ 300 + 1
        x = x Mod 300
        If x = 0 Then
            x = 300
            y = y - 1
        End If
        
        'Pogledamo kako obliko potrebuje register
        If registerFormat Then
            registerFormatPrint = Format(y, "00")
        Else
            registerFormatPrint = Format(y, "0")
        End If
        
        'Preverimo ali je toèka AO in jo damo v "A" podroèje, drugaèe v "B"
        If addressa >= CInt(zacAO) And addressa <= CInt(koncAO) Then
            If suffixEnable Then
                If Security Then
                    If Predpona = vbNullString Then
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_AR" & registerFormatPrint & "_DINT_" & obmocjeTEHNIK & "_DINT.F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_AR" & registerFormatPrint & "_REAL_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_AR" & registerFormatPrint & "_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                        End If
                    Else
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_AR" & registerFormatPrint & "_DINT_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_AR" & registerFormatPrint & "_REAL_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_AR" & registerFormatPrint & "_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                        End If
                    End If
                               
                Else
                    If Predpona = vbNullString Then
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_AR" & registerFormatPrint & "_DINT.F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_AR" & registerFormatPrint & "_REAL.F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_AR" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                        End If
                    Else
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_AR" & registerFormatPrint & "_DINT.F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_AR" & registerFormatPrint & "_REAL.F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_AR" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                        End If
                    End If
                End If
            Else
                If Security Then
                    If Predpona = vbNullString Then
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_AR" & registerFormatPrint & "_DINT_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_AR" & registerFormatPrint & "_REAL_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_AR" & registerFormatPrint & "_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                        End If
                    Else
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_AR" & registerFormatPrint & "_DINT_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_AR" & registerFormatPrint & "_REAL_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_AR" & registerFormatPrint & "_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                        End If
                    End If
                Else
                    If Predpona = vbNullString Then
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_AR" & registerFormatPrint & "_DINT.F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_AR" & registerFormatPrint & "_REAL.F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_AR" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                        End If
                    Else
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_AR" & registerFormatPrint & "_DINT.F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_AR" & registerFormatPrint & "_REAL.F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_AR" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                        End If
                    End If
                End If
            End If
        Else
            If suffixEnable Then
                If Security Then
                    If Predpona = vbNullString Then
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_AR" & registerFormatPrint & "_DINT_" & obmocjeTEHNOLOG & ".F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_AR" & registerFormatPrint & "_REAL_" & obmocjeTEHNOLOG & ".F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_AR" & registerFormatPrint & "_" & obmocjeTEHNOLOG & ".F_" & Format(x - 1, "000")
                        End If
                    Else
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_AR" & registerFormatPrint & "_DINT_" & obmocjeTEHNOLOG & ".F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_AR" & registerFormatPrint & "_REAL_" & obmocjeTEHNOLOG & ".F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_AR" & registerFormatPrint & "_" & obmocjeTEHNOLOG & ".F_" & Format(x - 1, "000")
                        End If
                    End If
                               
                Else
                    If Predpona = vbNullString Then
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_AR" & registerFormatPrint & "_DINT.F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_AR" & registerFormatPrint & "_REAL.F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_AR" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                        End If
                    Else
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_AR" & registerFormatPrint & "_DINT.F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_AR" & registerFormatPrint & "_REAL.F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_AR" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                        End If
                    End If
                End If
            Else
                If Security Then
                    If Predpona = vbNullString Then
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_AR" & registerFormatPrint & "_DINT_" & obmocjeTEHNOLOG & ".F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_AR" & registerFormatPrint & "_REAL_" & obmocjeTEHNOLOG & ".F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_AR" & registerFormatPrint & "_" & obmocjeTEHNOLOG & ".F_" & Format(x - 1, "000")
                        End If
                    Else
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_AR" & registerFormatPrint & "_DINT_" & obmocjeTEHNOLOG & ".F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_AR" & registerFormatPrint & "_REAL_" & obmocjeTEHNOLOG & ".F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_AR" & registerFormatPrint & "_" & obmocjeTEHNOLOG & ".F_" & Format(x - 1, "000")
                        End If
                    End If
                Else
                    If Predpona = vbNullString Then
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_AR" & registerFormatPrint & "_DINT.F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_AR" & registerFormatPrint & "_REAL.F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_AR" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                        End If
                    Else
                        If DataType Like "DINT" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_AR" & registerFormatPrint & "_DINT.F_" & Format(x - 1, "000")
                        ElseIf DataType Like "REAL" Then
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_AR" & registerFormatPrint & "_REAL.F_" & Format(x - 1, "000")
                        Else
                            Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_AR" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                        End If
                    End If
                End If
            End If
        End If
        
        Worksheets("tgd").Cells(cnt + 1, 1).Value = Worksheets(InputSheet).Cells(rw.Row, 1).Value
        Worksheets("tgd").Cells(cnt + 1, 3).Value = Worksheets(InputSheet).Cells(rw.Row, 3).Value
    End If
'M (do 2000)------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If InStr(Worksheets(InputSheet).Cells(rw.Row, 15).Value, "%M") > 0 And x < CInt(registerMcount) + 1 Then
        cnt = cnt + 1
        y = x \ 1000 + 1
        x = x Mod 1000
        If x = 0 Then
            x = 1000
            y = y - 1
        End If
        
        'Pogledamo kako obliko potrebuje register
        If registerFormat Then
            registerFormatPrint = Format(y, "00")
        Else
            registerFormatPrint = Format(y, "0")
        End If
        
        If suffixEnable Then
            If Security Then
                If Predpona = vbNullString Then
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_DR" & registerFormatPrint & "_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                Else
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_DR" & registerFormatPrint & "_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                End If
            Else
                If Predpona = vbNullString Then
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_DR" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                Else
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_DR" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                End If
            End If
        Else
            If Security Then
                If Predpona = vbNullString Then
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_DR" & registerFormatPrint & "_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                Else
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_DR" & registerFormatPrint & "_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                End If
            Else
                If Predpona = vbNullString Then
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_DR" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                Else
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_DR" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                End If
            End If
        End If
        Worksheets("tgd").Cells(cnt + 1, 1).Value = Worksheets(InputSheet).Cells(rw.Row, 1).Value
        Worksheets("tgd").Cells(cnt + 1, 3).Value = Worksheets(InputSheet).Cells(rw.Row, 3).Value
    End If
'Q (do 96)------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If InStr(Worksheets(InputSheet).Cells(rw.Row, 15).Value, "%Q") > 0 And x < CInt(registerQcount) + 1 Then
        cnt = cnt + 1
        y = x \ 1000 + 1
        x = x Mod 1000
        If x = 0 Then
            x = 1000
            y = y - 1
        End If
        
        'Pogledamo kako obliko potrebuje register. ZA DQ je drugacna
        If Not registerFormatDQ Then
            registerFormatPrint = "DRQ" & Format(y, "0")
        Else
            registerFormatPrint = "DRQ"
        End If
        
    'prilagojeno DRQ_A.F_CV (le en register)
        If suffixEnable Then
            If Security Then
                If Predpona = vbNullString Then
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_" & registerFormatPrint & "_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                Else
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_" & registerFormatPrint & "_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                End If
            Else
                If Predpona = vbNullString Then
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_0_" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                Else
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_0_" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                End If
            End If
        Else
            If Security Then
                If Predpona = vbNullString Then
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_" & registerFormatPrint & "_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                Else
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_" & registerFormatPrint & "_" & obmocjeTEHNIK & ".F_" & Format(x - 1, "000")
                End If
            Else
                If Predpona = vbNullString Then
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & PLCName & "_" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                Else
                    Worksheets("tgd").Cells(cnt + 1, 2).Value = "Fix32." & nodeName & "." & Predpona & "_" & PLCName & "_" & registerFormatPrint & ".F_" & Format(x - 1, "000")
                End If
            End If
        End If
        Worksheets("tgd").Cells(cnt + 1, 1).Value = Worksheets(InputSheet).Cells(rw.Row, 1).Value
        Worksheets("tgd").Cells(cnt + 1, 3).Value = Worksheets(InputSheet).Cells(rw.Row, 3).Value
    End If
Next rw

Worksheets("tgd").Cells(1, 1).Value = cnt
CSVname = InputBox("Vpiši ime CSV-ja (èe pustiš prazno/preklièeš, ga ne generira):")
If CSVname = vbNullString Then
    MsgBox "Prekinjeno!?"
    Exit Sub
Else
    Call saveAsCSVslo("tgd", CSVname, Environ("USERPROFILE") & "\Desktop\")
End If

End Sub


