Attribute VB_Name = "KreiranjePLCtock"
Global IndexM, odmik, Test_di, krmilnik
Sub ustvari_csv()
Dim message, Title, Default, krmilnik, ImeSistema, ImeTocke, Adresa, opis, Sistem1, Sistem2, Sistem3
Dim Index As Long, Index1 As Long, IndexM As Long, IndexR As Long, odmik As Long
Dim SteviloDI As Long, SteviloDO As Long, SteviloAQ As Long, SteviloAI As Long, SteviloPID As Long, SteviloTimer As Long, SteviloSis As Long
Dim SteviloRamp As Long
Dim PrviDI As Long, PrviDO As Long, PrviAI As Long, PrviAQ As Long, PrviPID As Long, PrviRamp As Long, PrviTimer As Long, PrviSis As Long
Dim Skaliranje1 As Long, Skaliranje10 As Long, PID As Long, Timer As Long, temporary As Long, MarkerE As Long, MarkerU As Long, MarkerALL As Long

Sheets("Sheet2").Select
Cells.Select
Selection.ClearContents
Range("A1").Select

Sheets("Sheet2").Cells(1, "A").Value = "Name"
Sheets("Sheet2").Cells(1, "B").Value = "DataType"
Sheets("Sheet2").Cells(1, "C").Value = "Description"
Sheets("Sheet2").Cells(1, "D").Value = "DataTypeID"
Sheets("Sheet2").Cells(1, "E").Value = "Retentive"
Sheets("Sheet2").Cells(1, "F").Value = "Force"
Sheets("Sheet2").Cells(1, "G").Value = "DisplayFormat"
Sheets("Sheet2").Cells(1, "H").Value = "ArrayDimension1"
Sheets("Sheet2").Cells(1, "I").Value = "ArrayDimension2"
Sheets("Sheet2").Cells(1, "J").Value = "Publish"
Sheets("Sheet2").Cells(1, "K").Value = "MarkAsUsed"
Sheets("Sheet2").Cells(1, "L").Value = "MaxLength"
Sheets("Sheet2").Cells(1, "M").Value = "InitialValue"
Sheets("Sheet2").Cells(1, "N").Value = "DataSource"
Sheets("Sheet2").Cells(1, "O").Value = "IOAddress"
Sheets("Sheet2").Cells(1, "P").Value = "IOAddressOffset"

'VNOS IMENA KRMILNIKA

krmilnik = Sheets("IOT").Cells(1, "I").Value

''Vnos PARAMETRA koliko %M imamo na krmilniku
'Message = "Vnesi število %M lokacij, ki so na krmilniku"
'Title = "Vnesi število %M lokacij, ki so na krmilniku"
'Default = "1024"
'MarkerALL = InputBox(Message, Title, Default)


'IZRAÈUN ŠTEVILA VHODOV, IZHODOV, REGULATORJEV, RAMP, TIMERJEV IN SISTEMOV
Index = 1
konec = 0
SteviloDI = 0
SteviloDO = 0
SteviloAI = 0
SteviloAQ = 0
SteviloPID = 0
SteviloRamp = 0
SteviloTimer = 0
SteviloSis = 0

Do Until konec = 1
    ImeTocke = Sheets("IOT").Cells(Index, "A").Value
    If ImeTocke = "" Then
        konec = 1
    Else
        konec = 0
        Index = Index + 1
    End If

    Test_di = Sheets("IOT").Cells(Index, "a").Value Like "*%I*"
        If Test_di = True Then
            SteviloDI = SteviloDI + 1
        End If

    Test_di = Sheets("IOT").Cells(Index, "a").Value Like "*%Q*"
        If Test_di = True Then
            SteviloDO = SteviloDO + 1
        End If

    Test_di = Sheets("IOT").Cells(Index, "a").Value Like "*%AQ*"
        If Test_di = True Then
            SteviloAQ = SteviloAQ + 1
        End If

    Test_di = Sheets("IOT").Cells(Index, "a").Value Like "*%AI*"
        If Test_di = True Then
            SteviloAI = SteviloAI + 1
        End If

    Test_di = Sheets("IOT").Cells(Index, "a").Value Like "*%PID*"
        If Test_di = True Then
            SteviloPID = SteviloPID + 1
        End If

   
    Test_di = Sheets("IOT").Cells(Index, "a").Value Like "*%TMR*"
        If Test_di = True Then
            SteviloTimer = SteviloTimer + 1
        End If
        
 Test_di = Sheets("IOT").Cells(Index, "a").Value Like "*%RAMP*"
        If Test_di = True Then
            SteviloRamp = SteviloRamp + 1
        End If
      
    Test_di = Sheets("IOT").Cells(Index, "a").Value Like "*%SIS*"
        If Test_di = True Then
              SteviloSis = SteviloSis + 1
        End If
        ime_tocke = ""

Loop

'KREIRANJE ZAÈETNIH NASLOVOV POSAMEZNE VRSTE SPREMENLJIVK
    PrviDI = Format(1, "00000")
    PrviDO = Format(PrviDI + SteviloDI, "00000")
    PrviAQ = Format(PrviDO + SteviloDO, "00000")
    PrviAI = Format(PrviAQ + SteviloAQ, "00000")
    PrviPID = Format(PrviAI + SteviloAI, "00000")
    PrviRamp = Format(PrviPID + SteviloPID, "00000")
    PrviTimer = Format(PrviRamp + SteviloRamp, "00000")
    PrviSis = Format(PrviTimer + SteviloTimer, "00000")
    
odmik = 2

'KREIRANJE SUROVIH DIGITALNIH VHODOV
index_tipa = 1
    y = DI("I", index_tipa, PrviDI, SteviloDI, odmik, "", "")
index_tipa = index_tipa + 1

'KREIRANJE MARKERJEV DIGITALNIH VHODOV
IndexM = 1
  y = DI("M", IndexM, PrviDI, SteviloDI, odmik, "", "")

'KREIRANJE MARKERJEV ANALOGNIH VHODOV ZA SIGNAL IZPAD SENZORJA
  y = DI("M", IndexM, PrviAI, SteviloAI, odmik, "_E_SENS", ", izp. tip.")

'KREIRANJE MARKERJEV ANALOGNIH VHODOV ZA SIGNAL ALARM
  y = DI("M", IndexM, PrviAI, SteviloAI, odmik, "_A_HIHI", ", alarm")
  y = DI("M", IndexM, PrviAI, SteviloAI, odmik, "_A_HI", ", opozorilo")
  y = DI("M", IndexM, PrviAI, SteviloAI, odmik, "_A_LO", ", opozorilo")
  y = DI("M", IndexM, PrviAI, SteviloAI, odmik, "_A_LOLO", ", alarm")

'KREIRANJE MARKERJEV ANALOGNIH VHODOV ZA SIGNAL OMOGOÈI ALARMIRANJE
  y = DI("M", IndexM, PrviAI, SteviloAI, odmik, "_A_EN", ", enable")

'KREIRANJE MARKERJEV ANALOGNIH VHODOV ZA KVIT ALARMA
  y = DI("M", IndexM, PrviAI, SteviloAI, odmik, "_KVIT", ", kvit")
  
'KREIRANJE SPREMENLJIVK MARKERJEV PID REGULATORJEV
  y = DI("M", IndexM, PrviPID, SteviloPID, odmik, "_MAN", ", manual")
  y = DI("M", IndexM, PrviPID, SteviloPID, odmik, "_ROC", ", rocno")

'KREIRANJE SPREMENLJIVK MARKERJEV ZA SISTEME (KLIMATE)
y = DI("M", IndexM, PrviSis, SteviloSis, odmik, "_AUTO", "Avtomatsko")
y = DI("M", IndexM, PrviSis, SteviloSis, odmik, "_ROCNO", "Rocno")
y = DI("M", IndexM, PrviSis, SteviloSis, odmik, "_SERVIS", "Servis")
y = DI("M", IndexM, PrviSis, SteviloSis, odmik, "_VKLOP_SCADA", "Vklop Scada")
y = DI("M", IndexM, PrviSis, SteviloSis, odmik, "_KVIT_SCADA", "Kvit Scada")
y = DI("M", IndexM, PrviSis, SteviloSis, odmik, "_DEL_A", "Del. motnja")
                
'PARAMETER ALHILO ENABLE
  y = DI("M", IndexM, PrviSis, SteviloSis, odmik, "_AL_ENABLE_TH", "Alarm. temp. vlag. enable")
  y = DI("M", IndexM, PrviSis, SteviloSis, odmik, "_AL_ENABLE_P", "Alarm. tlaka enable")

'KREIRANJE MARKERJEV ZA SPLOŠNE NAPAKE
'VNOS PARAMETROV MARKERJI NAPAK
message = "Predvidena prva lokacija za markerje NAPAK. Po želji lahko spremeniš."
Title = "Predvidena prva lokacija za markerje NAPAK"
MarkerE = (Application.WorksheetFunction.RoundDown((IndexM / 16), 0) * 16) + 1
If MarkerE <= IndexM Then
    MarkerE = MarkerE + 16
End If
Default = MarkerE
MarkerE = InputBox(message, Title, Default)

IndexM = MarkerE
  y = E(IndexM, krmilnik, odmik, "E_POZAR", "Pozar")
  y = E(IndexM, krmilnik, odmik, "E_ZAS_IZKL", "Zasilni izklop")
  y = E(IndexM, krmilnik, odmik, "E_IZP_ZASC", "Izpad zascit")
  y = E(IndexM, krmilnik, odmik, "E_IZP_KRM", "Izpad krmilne napetosti")
  y = E(IndexM, krmilnik, odmik, "E_IZP_NAP", "Izpad mrezne napetosti")
  y = E(IndexM, krmilnik, odmik, "E_ROCNO", "Rocni vklop")
  y = E(IndexM, krmilnik, odmik, "E_PLC_BAT", "Prazna baterija CPU")
  y = E(IndexM, krmilnik, odmik, "E_HRD_CPU", "Izpad CPU")
  y = E(IndexM, krmilnik, odmik, "E_LOS_IOM", "Izpad I/O modula")
  'dodano BaM 22.06.2023
  y = E(IndexM, krmilnik, odmik, "E_OVR_TMP", "Pregrevanje CPU enote")
  y = E(IndexM, krmilnik, odmik, "E_CFG_MM", "HRDW konfiguracija napaka")
  y = E(IndexM, krmilnik, odmik, "E_NO_PROG", "Ni programa")
  y = E(IndexM, krmilnik, odmik, "E_FORCE", "Prisoten FORCE na krmilniku")
  y = E(IndexM, krmilnik, odmik, "E_OV_SWP", "Prekoracen Sweep time")
  y = E(IndexM, krmilnik, odmik, "E_BAD_RAM", "Možnost izgube podatkov v RAM-u")
  y = E(IndexM, krmilnik, odmik, "E_ETH_LINK", "Eth. mreža ni dosegljiva")
  y = E(IndexM, krmilnik, odmik, "E_ETH_INTERFACE", "Napaka ethernet kartice")
  y = E(IndexM, krmilnik, odmik, "E_REZ_SIS_01", "Napaka sistemska 01")
  y = E(IndexM, krmilnik, odmik, "E_REZ_SIS_02", "Napaka sistemska 02")
  y = E(IndexM, krmilnik, odmik, "E_REZ_SIS_03", "Napaka sistemska 03")
  y = E(IndexM, krmilnik, odmik, "E_REZ_SIS_04", "Napaka sistemska 04")
  y = E(IndexM, krmilnik, odmik, "E_REZ_SIS_05", "Napaka sistemska 05")
  y = E(IndexM, krmilnik, odmik, "E_REZ_SIS_06", "Napaka sistemska 06")
  y = E(IndexM, krmilnik, odmik, "E_REZ_SIS_07", "Napaka sistemska 07")
  y = E(IndexM, krmilnik, odmik, "E_REZ_SIS_08", "Napaka sistemska 08")
  y = E(IndexM, krmilnik, odmik, "E_REZ_SIS_09", "Napaka sistemska 09")
  y = E(IndexM, krmilnik, odmik, "E_REZ_SIS_10", "Napaka sistemska 10")
  y = E(IndexM, krmilnik, odmik, "E_REZ_SIS_11", "Napaka sistemska 11")
  y = E(IndexM, krmilnik, odmik, "E_REZ_SIS_12", "Napaka sistemska 12")
  y = E(IndexM, krmilnik, odmik, "E_REZ_SIS_13", "Napaka sistemska 13")
  y = E(IndexM, krmilnik, odmik, "E_REZ_SIS_14", "Napaka sistemska 14")
  y = E(IndexM, krmilnik, odmik, "E_REZ_SIS_15", "Napaka sistemska 15")

  
'POSTAVLJANJE INDEKSA M NA CELO ŠESTNAJSTICO + 1 ZARADI ADRESNEGA POLJA V CME
IndexM = MarkerE + 32

'DODAMO BITE ZA STATUS MODULOV PLC_VD_SLOT_X_Y_E ... MPSPK3: Napaka RACK 00 Slot 00
Adresa = "%M" & Format(IndexM, "00000")
For rack = 0 To 3
    If rack = 0 Then
        For slot = 0 To 16
            Adresa = "%M" & Format(IndexM, "00000")
            ImeTocke = krmilnik & "_VD_SLOT_" & Format(rack, "00") & "_" & Format(slot, "00") & "_E"
            opis = krmilnik & ": " & "Napaka RACK " & Format(rack, "00") & " Slot " & Format(slot, "00")
            
            Sheets("Sheet2").Cells(odmik, "A").Value = ImeTocke
            Sheets("Sheet2").Cells(odmik, "B").Value = "BOOL"
            Sheets("Sheet2").Cells(odmik, "C").Value = opis
            Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
            
            IndexM = IndexM + 1
            odmik = odmik + 1
            
        Next
    End If
    If rack <> 0 Then
        For slot = 0 To 10
            Adresa = "%M" & Format(IndexM, "00000")
            ImeTocke = krmilnik & "_VD_SLOT_" & Format(rack, "00") & "_" & Format(slot, "00") & "_E"
            opis = krmilnik & ": " & "Napaka RACK " & Format(rack, "00") & " Slot " & Format(slot, "00")
            
            Sheets("Sheet2").Cells(odmik, "A").Value = ImeTocke
            Sheets("Sheet2").Cells(odmik, "B").Value = "BOOL"
            Sheets("Sheet2").Cells(odmik, "C").Value = opis
            Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
            
            IndexM = IndexM + 1
            odmik = odmik + 1
            
        Next
    End If
Next

'KREIRANJE MARKERJEV ZA NAPAKE NA SISTEMIH (MOTORJI, ÈRPALKE, FILTRI...)
Dim zacadr() As Variant
ReDim zacadr(SteviloSis, 1)
Index = 1
'poveèamo za 4 ker smo drugaèe not bit aligned
IndexM = IndexM + 14

Do Until Index = SteviloSis + 1

'PARAMETER GENERALNA NAPAKA
        Adresa = "%M" & Format(IndexM, "00000")
        ImeSistema = Sheets("IOT").Cells(Index + PrviSis, "B").Value
        ImeTocke = Sheets("IOT").Cells(Index + PrviSis, "C").Value
        opis = Sheets("IOT").Cells(Index + PrviSis, "D").Value
        
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeTocke & "_VD_E_GEN"
        Sheets("Sheet2").Cells(odmik, "B").Value = "BOOL"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": Izpad sistema"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        
        'Zapis naslovov za vsak sistem
        zacadr(Index - 1, 0) = ImeSistema
        zacadr(Index - 1, 1) = IndexM + 1
        
        Index = Index + 1
        IndexM = IndexM + 32
        odmik = odmik + 1
Loop
        'Zapis za preostale
        zacadr(Index - 1, 0) = "OSTALO"
        zacadr(Index - 1, 1) = IndexM
        Dim Adresa_tmp As Integer
        Adresa_tmp = 0
'NAPAKE ÈRPALK IN MOTORJEV
Index = 1
Do Until Index = SteviloDO + 1
    Adresa = "%M" & Format(IndexM, "00000")
    ImeSistema = Sheets("IOT").Cells(Index + PrviDO, "B").Value
    ImeTocke = Sheets("IOT").Cells(Index + PrviDO, "C").Value
    opis = Sheets("IOT").Cells(Index + PrviDO, "D").Value

    If (Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*F*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*P*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*VM*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*Ve*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*SIC*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*SI*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*EG*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Elektrièni*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*ventilator*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*èrpalka*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*crpalka*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Èrpalka*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Ogrevalne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*ogrevalne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Ventilator*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*vent*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Vent*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*crp*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Crp*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*FFU*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Grelne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*grelne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Rešetke*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*rešetke*")) Then
            Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VD_" & ImeTocke & "_E_DEL"
            Sheets("Sheet2").Cells(odmik, "B").Value = "BOOL"
            Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", nap. vkl."
            
                For i = 0 To SteviloSis
                    If zacadr(i, 0) = ImeSistema Then
                        IndexM = zacadr(i, 1)
                        zacadr(i, 1) = zacadr(i, 1) + 1
                        Adresa_tmp = 1
                        Exit For
                    End If
                Next i

                If Adresa_tmp <> 0 Then
                    Adresa = "%M" & Format(IndexM, "00000")
                    Adresa_tmp = 0
                Else
                    IndexM = zacadr(i - 1, 1)
                    zacadr(i - 1, 1) = zacadr(i - 1, 1) + 1
                    Adresa = "%M" & Format(IndexM, "00000")
                End If
            
            Sheets("Sheet2").Cells(odmik, "O").Value = Adresa

            IndexM = IndexM + 1
            odmik = odmik + 1
    End If
    Adresa = "%M" & Format(IndexM, "00000")
    If ((Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*FY*") Or _
         Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*PY*") Or _
         Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*SIC*")) _
        And _
        (Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Ventilator*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*vent*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Vent*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*crp*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Crp*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*ventilator*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*èrpalka*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*crpalka*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Èrpalka*"))) Then
            Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VD_" & ImeTocke & "_E_FP"
            Sheets("Sheet2").Cells(odmik, "B").Value = "BOOL"
            Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", nap. f.p."
            
                For i = 0 To SteviloSis
                    If zacadr(i, 0) = ImeSistema Then
                        IndexM = zacadr(i, 1)
                        zacadr(i, 1) = zacadr(i, 1) + 1
                        Adresa_tmp = 1
                        Exit For
                    End If
                Next i

                If Adresa_tmp <> 0 Then
                    Adresa = "%M" & Format(IndexM, "00000")
                    Adresa_tmp = 0
                Else
                    IndexM = zacadr(i - 1, 1)
                    zacadr(i - 1, 1) = zacadr(i - 1, 1) + 1
                    Adresa = "%M" & Format(IndexM, "00000")
                End If
            
            Sheets("Sheet2").Cells(odmik, "O").Value = Adresa

            IndexM = IndexM + 1
            odmik = odmik + 1
    End If

    Index = Index + 1
Loop
            
'ZAMAŠENOST FILTROV
Index = 1
Do Until Index = SteviloDI + 1
    Adresa = "%M" & Format(IndexM, "00000")
    ImeSistema = Sheets("IOT").Cells(Index + PrviDI, "B").Value
    ImeTocke = Sheets("IOT").Cells(Index + PrviDI, "C").Value
    opis = Sheets("IOT").Cells(Index + PrviDI, "D").Value
               
    If Sheets("IOT").Cells(Index + PrviDI, "C").Value Like ("*PDSA*") Then
        
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VD_" & ImeTocke & "_E_FILT"
        Sheets("Sheet2").Cells(odmik, "B").Value = "BOOL"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis
        
                For i = 0 To SteviloSis
                    If zacadr(i, 0) = ImeSistema Then
                        IndexM = zacadr(i, 1)
                        zacadr(i, 1) = zacadr(i, 1) + 1
                        Adresa_tmp = 1
                        Exit For
                    End If
                Next i

                If Adresa_tmp <> 0 Then
                    Adresa = "%M" & Format(IndexM, "00000")
                    Adresa_tmp = 0
                Else
                    IndexM = zacadr(i - 1, 1)
                    zacadr(i - 1, 1) = zacadr(i - 1, 1) + 1
                    Adresa = "%M" & Format(IndexM, "00000")
                End If
        
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        
        IndexM = IndexM + 1
        odmik = odmik + 1
    End If
    Index = Index + 1
Loop
        
IndexM = zacadr(SteviloSis, 1) + 16
       
'VKLOPI ZA VENTILATORJE IN ÈRPALKE
            
'RESET OBRATOVALNIUH UR IN ŠTEVILO VKLOPOV
Index = 1
Do Until Index = SteviloDO + 1
    'Biti za RESET OBRH in reset ST_VKL
    ImeSistema = Sheets("IOT").Cells(Index + PrviDO, "B").Value
    ImeTocke = Sheets("IOT").Cells(Index + PrviDO, "C").Value
    opis = Sheets("IOT").Cells(Index + PrviDO, "D").Value
    
    If (Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*F*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*P*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*VM*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*Ve*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*SIC*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*SI*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*EG*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Elektrièni*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*ventilator*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*èrpalka*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*crpalka*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Èrpalka*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Ogrevalne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*ogrevalne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Ventilator*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*vent*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Vent*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*crp*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Crp*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*FFU*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Grelne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*grelne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Rešetke*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*rešetke*")) Then
       
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VD_" & ImeTocke & "_OBRHD_R"
        Sheets("Sheet2").Cells(odmik, "B").Value = "BOOL"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", res. OBRH"
        Sheets("Sheet2").Cells(odmik, "O").Value = "%M" & Format(IndexM, "00000")
        IndexM = IndexM + 1
        odmik = odmik + 1
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VD_" & ImeTocke & "_ST_VKL_R"
        Sheets("Sheet2").Cells(odmik, "B").Value = "BOOL"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", res. št. vkl."
        Sheets("Sheet2").Cells(odmik, "O").Value = "%M" & Format(IndexM, "00000")
        IndexM = IndexM + 1
        odmik = odmik + 1
    End If
    Index = Index + 1
Loop




'KREIRANJE SPREMENLJIVK ZA SERVIS----------------------------------------------------------------------------------
message = "Predvidena prva lokacija za markerje DO in DI - _OFF, _AU, _MN, _SR, _BL, _INV. Po želji lahko spremeniš."
Title = "Predvidena prva lokacija za markerje _OFF, _AU, _MN, _SR, _BL, _INV."
MarkerE = 2001
If MarkerE <= IndexM Then
    MarkerE = MarkerE + 16
End If
Default = MarkerE

MarkerE = InputBox(message, Title, Default)

IndexM = MarkerE
    
'KREIRAMO _SB bite za DOje-----------------------------------------------
y = DI("M", IndexM, PrviDI, SteviloDI, odmik, "_SB", ", servis bit")

'KREIRAMO _SV bite za DOje-----------------------------------------------
y = DI("M", IndexM, PrviDI, SteviloDI, odmik, "_SV", ", vrednost v serv.")
    
'KREIRAMO _OFF bite za DOje-----------------------------------------------
y = DI("M", IndexM, PrviDO, SteviloDO, odmik, "_OFF", ", vrednost v OFF")
  
'KREIRAMO _AU bite za DOje-----------------------------------------------
y = DI("M", IndexM, PrviDO, SteviloDO, odmik, "_AU", ", vrednost v AUTO")

'KREIRAMO _MN bite za DOje-----------------------------------------------
y = DI("M", IndexM, PrviDO, SteviloDO, odmik, "_MN", ", vrednost v ROCNO")

'KREIRAMO _SR bite za DOje-----------------------------------------------
y = DI("M", IndexM, PrviDO, SteviloDO, odmik, "_SR", ", vrednost v SERVIS")

'KREIRAMO _BL bite za DOje-----------------------------------------------*
y = DI("M", IndexM, PrviDO, SteviloDO, odmik, "_BL", ", INTERLOK")

'------------------------------AO-----------------------------------------AO toèke imajo samo BL od VD vrednosti
IndexM = IndexM
Index = 1
Do Until Index = SteviloAQ + 1
    Adresa = "%M" & Format(IndexM, "00000")
    ImeSistema = Sheets("IOT").Cells(Index + PrviAQ, "B").Value
    ImeTocke = Sheets("IOT").Cells(Index + PrviAQ, "C").Value
    opis = Sheets("IOT").Cells(Index + PrviAQ, "D").Value
    
    For innerIndex = 1 To SteviloDO
        Duplicate = False
        ImeTockeDO = Sheets("IOT").Cells(innerIndex + PrviDO, "C").Value
        If ImeTocke Like ImeTockeDO Then
            Duplicate = True
            Exit For
        End If
    Next
    If Duplicate Then
        Index = Index + 1
    Else
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VD_" & ImeTocke & "_BL"
        Sheets("Sheet2").Cells(odmik, "B").Value = "BOOL"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", INTERLOK"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexM = IndexM + 1
        odmik = odmik + 1
        Index = Index + 1
    End If
Loop


'KREIRAMO _INV bite za DOje-----------------------------------------------
y = DI("M", IndexM, PrviDO, SteviloDO, odmik, "_INV", ", Inverse")


'KREIRANJE DIGITALNIH IZHODOV
    index_tipa = 1
    y = DI("Q", index_tipa, PrviDO, SteviloDO, odmik, "", "")
          
'KREIRANJE ANALOGNIH IZHODOV
    index_tipa = 1
    y = DI("AQ", index_tipa, PrviAQ, SteviloAQ, odmik, "", "")
    
'KREIRANJE SUROVE VREDNOSTI ANALOGNIH VHODOV
    index_tipa = 1
    y = DI("AI", index_tipa, PrviAI, SteviloAI, odmik, "", "")
    index_tipa = index_tipa + 1
    
'KREIRANJE INŽENIRSKE VREDNOSTI ANALOGNIH VHODOV
    IndexR = 1
    y = DI("R", IndexR, PrviAI, SteviloAI, odmik, "", "")
    
'KREIRANJE INŽENIRSKE VREDNOSTI ANALOGNIH IZHODOV
    y = DI("R", IndexR, PrviAQ, SteviloAQ, odmik, "", "")
        
'KREIRANJE MEJ ALARMIRANJA ANALOGNIH VHODOV
    y = DI("R", IndexR, PrviAI, SteviloAI, odmik, "_HIHI", ", meja")
    y = DI("R", IndexR, PrviAI, SteviloAI, odmik, "_HI", ", meja")
    y = DI("R", IndexR, PrviAI, SteviloAI, odmik, "_LO", ", meja")
    y = DI("R", IndexR, PrviAI, SteviloAI, odmik, "_LOLO", ", meja")
  
'POSTAVLJANJE INDEKSA R NA CELO DESETICO + 1 ZARADI ADRESNEGA POLJA V CME
Index = IndexR / 10
IndexR = (Index + 1) * 10 + 1

'KREIRANJE SPREMENLJIVK PID REGULATORJEV
Index = 1
Do Until Index = SteviloPID + 1
        
    'PARAMETER KP
        Adresa = "%R" & Format(IndexR, "00000")
        ImeSistema = Sheets("IOT").Cells(Index + PrviPID, "B").Value
        ImeTocke = Sheets("IOT").Cells(Index + PrviPID, "C").Value
        opis = Sheets("IOT").Cells(Index + PrviPID, "D").Value
                
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_KP"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", KP"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
      
        IndexR = IndexR + 1
        odmik = odmik + 1
      
   'PARAMETER KD
        Adresa = "%R" & Format(IndexR, "00000")
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_KD"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", KD"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
      
        IndexR = IndexR + 1
        odmik = odmik + 1
    
    'PARAMETER KI
        Adresa = "%R" & Format(IndexR, "00000")
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_KI"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", KI"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
      
        IndexR = IndexR + 1
        odmik = odmik + 1
      
    'PARAMETER CVB
        Adresa = "%R" & Format(IndexR, "00000")
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_CVB"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", CV bias"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
      
        IndexR = IndexR + 1
        odmik = odmik + 1
      
    'PARAMETER UC
        Adresa = "%R" & Format(IndexR, "00000")
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_UC"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", UC"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
      
        IndexR = IndexR + 1
        odmik = odmik + 1
      
    'PARAMETER LC
        Adresa = "%R" & Format(IndexR, "00000")
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_LC"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", LC"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
      
        IndexR = IndexR + 1
        odmik = odmik + 1
      
    'PARAMETER CV
        Adresa = "%R" & Format(IndexR, "00000")
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_CV"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", CV"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
      
        IndexR = IndexR + 1
        odmik = odmik + 1
      
    'PARAMETER OPV
        Adresa = "%R" & Format(IndexR, "00000")
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_OPV"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", OPV"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
      
        IndexR = IndexR + 1
        odmik = odmik + 1
      
      'PARAMETER SP
        Adresa = "%R" & Format(IndexR, "00000")
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_SP"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", SP"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
      
        IndexR = IndexR + 1
        odmik = odmik + 1
      
    'PARAMETER SP DEJ
        Adresa = "%R" & Format(IndexR, "00000")
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_SP_DEJ"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", SP dej."
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
            
        Index = Index + 1
        IndexR = IndexR + 1
        odmik = odmik + 1

Loop
Index = IndexR / 10
IndexR = (Index + 1) * 10 + 1
'KREIRANJE RAMP
Index = 1
Do Until Index = SteviloRamp + 1
     
     'KREIRANJE SPREMENLJIVKE YPRIXMIN
        Adresa = "%R" & Format(IndexR, "00000")
        ImeSistema = Sheets("IOT").Cells(Index + PrviRamp, "B").Value
        ImeTocke = Sheets("IOT").Cells(Index + PrviRamp, "C").Value
        opis = Sheets("IOT").Cells(Index + PrviRamp, "D").Value
        
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_YPRIXMIN"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", Y pri Xmin"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
 
        IndexR = IndexR + 1
        odmik = odmik + 1
      
     'KREIRANJE SPREMENLJIVKE YPRIXMAX
        Adresa = "%R" & Format(IndexR, "00000")
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_YPRIXMAX"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", Y pri Xmax"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
 
        IndexR = IndexR + 1
        odmik = odmik + 1
       
     'KREIRANJE SPREMENLJIVKE XMIN
        Adresa = "%R" & Format(IndexR, "00000")
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_XMIN"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", Xmin"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
 
        IndexR = IndexR + 1
        odmik = odmik + 1
      
    'KREIRANJE SPREMENLJIVKE XMAX
        Adresa = "%R" & Format(IndexR, "00000")
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_XMAX"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", Xmax"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
 
        IndexR = IndexR + 1
        odmik = odmik + 1
      
      'KREIRANJE SPREMENLJIVKE IZHOD IZ RAMPE
        Adresa = "%R" & Format(IndexR, "00000")
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", out"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
       
        Index = Index + 1
        IndexR = IndexR + 1
        odmik = odmik + 1
      
Loop

'KREIRANJE OFF variabel za DO-je in AQ-je
 '------------------------------AQ--------------------------------------------
message = "Predvidena prva lokacija za registre OFF, AUTO, MANUAL, SERVICE (_OFF, _AU, _MN, _SR) za AQ-variable. Po želji lahko spremeniš."
Title = "Predvidena prva lokacija za registre AQ[_OFF, _AU, _MN, _SR]"
MarkerE = IndexR
Default = MarkerE
MarkerE = InputBox(message, Title, Default)

IndexR = MarkerE
    
'KREIRAMO _OFF za AQje-------------------------------------------------------------------------
y = DI("R", IndexR, PrviAQ, SteviloAQ, odmik, "_OFF", ", vrednost v OFF")

'KREIRAMO _AU za AQje--------------------------------------------------------------------------
y = DI("R", IndexR, PrviAQ, SteviloAQ, odmik, "_AU", ", vrednost v AUTO")

'KREIRAMO _MN za AQje--------------------------------------------------------------------------
y = DI("R", IndexR, PrviAQ, SteviloAQ, odmik, "_MN", ", vrednost v ROCNO")

'KREIRAMO _SR za AQje-------------------------------------------------------------------------
y = DI("R", IndexR, PrviAQ, SteviloAQ, odmik, "_SR", ", vrednost v SERVIS")

'KREIRANJE SPREMENLJIVK OBRATOVALNE URE SISTEMOV
''VNOS PARAMETROV SKALIRANJE 1:10 DINT
message = "Predviden prvi register skaliranja 1:10 DINT. Po želji lahko spremeniš."
Title = "Predviden prvi register skaliranja 1:10 DINT"
Skaliranje10DINT = (Application.WorksheetFunction.RoundDown((IndexR / 300), 0) * 300) + 1
If Skaliranje10DINT <= IndexR Then
    Skaliranje10DINT = Skaliranje10DINT + 300
End If
Default = Skaliranje10DINT
Skaliranje10DINT = InputBox(message, Title, Default)
Index = 1
IndexR = Skaliranje10DINT
    y = DI("RD", IndexR, PrviSis, SteviloSis, odmik, "_OBRHD", "Obrat. ure sistema")
  
    
'OBRATOVALNE URE ZA VENTILATORJE IN ÈRPALKE (Krt)
Index = 1
Do Until Index = SteviloDO + 1

    Adresa = "%R" & Format(IndexR, "00000")
    ImeSistema = Sheets("IOT").Cells(Index + PrviDO, "B").Value
    ImeTocke = Sheets("IOT").Cells(Index + PrviDO, "C").Value
    opis = Sheets("IOT").Cells(Index + PrviDO, "D").Value
    
    If (Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*F*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*P*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*VM*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*Ve*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*SIC*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*SI*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*EG*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Elektrièni*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*ventilator*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*èrpalka*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*crpalka*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Èrpalka*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Ogrevalne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*ogrevalne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Ventilator*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*vent*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Vent*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*crp*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Crp*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*FFU*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Grelne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*grelne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Rešetke*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*rešetke*")) Then
            Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_OBRHD"
            Sheets("Sheet2").Cells(odmik, "B").Value = "DINT"
            Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", obr. ure"
            Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
            
            IndexR = IndexR + 2
            odmik = odmik + 1
    End If
        Index = Index + 1
Loop


''VNOS PARAMETROV SKALIRANJE 1:10
message = "Predviden prvi register skaliranja 1:10 registrov za OBRH. Po želji lahko spremeniš."
Title = "Predviden prvi register skaliranja 1:10"
Skaliranje10 = (Application.WorksheetFunction.RoundDown((IndexR / 300), 0) * 300) + 1
If Skaliranje10 <= IndexR Then
    Skaliranje10 = Skaliranje10DINT + 300
End If
Default = Skaliranje10
Skaliranje10 = InputBox(message, Title, Default)
IndexR = Skaliranje10

'VREDNOSTI TIPAL TLAKOV ZA SCADO
Index = 1
Do Until Index = SteviloAI + 1
        ImeSistema = Sheets("IOT").Cells(Index + PrviAI, "B").Value
        ImeTocke = Sheets("IOT").Cells(Index + PrviAI, "C").Value
        opis = Sheets("IOT").Cells(Index + PrviAI, "D").Value
                
        
            If Sheets("IOT").Cells(Index + PrviAI, "C").Value Like ("*PIC*") And _
               Sheets("IOT").Cells(Index + PrviAI, "D").Value Like ("*Kanal*") Then
                Adresa = "%R" & Format(IndexR, "00000")
                Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_SC"
                Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
                Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", scada"
                Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
                IndexR = IndexR + 1
                odmik = odmik + 1
            
                Adresa = "%R" & Format(IndexR, "00000")
                Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_SP_SC"
                Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
                Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", scada"
                Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
                IndexR = IndexR + 1
                odmik = odmik + 1
            
                Adresa = "%R" & Format(IndexR, "00000")
                Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_HIHI_SC"
                Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
                Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", meja"
                Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
                IndexR = IndexR + 1
                odmik = odmik + 1
                
                Adresa = "%R" & Format(IndexR, "00000")
                Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_HI_SC"
                Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
                Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", meja"
                Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
                IndexR = IndexR + 1
                odmik = odmik + 1
                
                Adresa = "%R" & Format(IndexR, "00000")
                Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_LO_SC"
                Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
                Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", meja"
                Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
                IndexR = IndexR + 1
                odmik = odmik + 1
            
                Adresa = "%R" & Format(IndexR, "00000")
                Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_LOLO_SC"
                Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
                Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", meja"
                Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
                IndexR = IndexR + 1
                odmik = odmik + 1
            
            End If
    Index = Index + 1
Loop
 
'KREIRANJE ZAKASNITEV ALARMIRANJA ANALOGNIH VHODOV
''VNOS PARAMETROV SKALIRANJE 1:1
message = "Predviden prvi register skaliranja 1:1. Po želji lahko spremeniš."
Title = "Predviden prvega registra skaliranja 1:1"
Skaliranje1 = (Application.WorksheetFunction.RoundDown((IndexR / 300), 0) * 300) + 1
If Skaliranje1 <= IndexR Then
    Skaliranje1 = Skaliranje1 + 300
End If
Default = Skaliranje1
Skaliranje1 = InputBox(message, Title, Default)

IndexR = Skaliranje1
  y = DI("R", IndexR, PrviAI, SteviloAI, odmik, "_ZAK1", ", zak. opoz.")
  y = DI("R", IndexR, PrviAI, SteviloAI, odmik, "_ZAK2", ", zak. alarm.")
           
'KREIRANJE SPREMENLJIVK OBRATOVALNI DNEVI, STANJE SISTEMA IN ZAKASNITEV ALARMIRANJA OB ZAGONU
  y = DI("R", IndexR, PrviSis, SteviloSis, odmik, "_SIS_STATUS", "Stanje sistema")
'KREIRANJE SPREMENLJIVK OBRATOVALNI DNEVI, STANJE SISTEMA IN ZAKASNITEV ALARMIRANJA OB ZAGONU
  y = DI("R", IndexR, PrviSis, SteviloSis, odmik, "_VODENJE", "Režim sistema - 1:AVTO, 2:ROCNO, 3:SERVIS")
  
'KREIRAMO _RZ za DOje-------------------------------------------------------------------------
y = DI("R", IndexR, PrviDO, SteviloDO, odmik, "_RZ", ", Rezim")

'-----------------------AO------------------------gremo še æez AO toèke in pazimo da ne podvajamo tagov
Index = 1
Do Until Index = SteviloAQ + 1
    Adresa = "%R" & Format(IndexR, "00000")
    ImeSistema = Sheets("IOT").Cells(Index + PrviAQ, "B").Value
    ImeTocke = Sheets("IOT").Cells(Index + PrviAQ, "C").Value
    opis = Sheets("IOT").Cells(Index + PrviAQ, "D").Value
    
    For innerIndex = 1 To SteviloDO
        Duplicate = False
        ImeTockeDO = Sheets("IOT").Cells(innerIndex + PrviDO, "C").Value
        If ImeTocke Like ImeTockeDO Then
            Duplicate = True
            Exit For
        End If
    Next
    If Duplicate Then
        Index = Index + 1
    Else
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_RZ"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", Rezim"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
        Index = Index + 1
    End If
Loop

'KREIRAMO _RZ za DOje-------------------------------------------------------------------------
y = DI("R", IndexR, PrviDO, SteviloDO, odmik, "_S", ", Status")
'-----------------------AO------------------------gremo še æez AO toèke in pazimo da ne podvajamo tagov
Index = 1
Do Until Index = SteviloAQ + 1
    Adresa = "%R" & Format(IndexR, "00000")
    ImeSistema = Sheets("IOT").Cells(Index + PrviAQ, "B").Value
    ImeTocke = Sheets("IOT").Cells(Index + PrviAQ, "C").Value
    opis = Sheets("IOT").Cells(Index + PrviAQ, "D").Value
    
    For innerIndex = 1 To SteviloDO
        Duplicate = False
        ImeTockeDO = Sheets("IOT").Cells(innerIndex + PrviDO, "C").Value
        If ImeTocke Like ImeTockeDO Then
            Duplicate = True
            Exit For
        End If
    Next
    If Duplicate Then
        Index = Index + 1
    Else
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_S"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", Status"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
        Index = Index + 1
    End If
Loop

'KREIRANJE ST_VKL
''VNOS PARAMETROV SKALIRANJE 1:1 za DINT registre
message = "Predviden prvi register skaliranja 1:1 za DINT registre. Po želji lahko spremeniš."
Title = "Predviden prvega registra skaliranja 1:1 DINT registrov"
Skaliranje1DINT = (Application.WorksheetFunction.RoundDown((IndexR / 300), 0) * 300) + 1
If Skaliranje1DINT <= IndexR Then
    Skaliranje1DINT = Skaliranje1DINT + 300
End If
Default = Skaliranje1DINT
Skaliranje1DINT = InputBox(message, Title, Default)

IndexR = Skaliranje1DINT
 
 
'OBRATOVALNI DNEVI ZA VENTILATORJE IN ÈRPALKE (Krt)
Index = 1
Do Until Index = SteviloDO + 1

    'VKLOP
    Adresa = "%R" & Format(IndexR, "00000")
    ImeSistema = Sheets("IOT").Cells(Index + PrviDO, "B").Value
    ImeTocke = Sheets("IOT").Cells(Index + PrviDO, "C").Value
    opis = Sheets("IOT").Cells(Index + PrviDO, "D").Value
            
    If (Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*F*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*P*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*VM*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*Ve*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*SIC*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*SI*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "C").Value Like ("*EG*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Elektrièni*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*ventilator*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*èrpalka*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*crpalka*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Èrpalka*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Ogrevalne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*ogrevalne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Ventilator*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*vent*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Vent*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*crp*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Crp*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*FFU*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Grelne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*grelne*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*Rešetke*") Or _
        Sheets("IOT").Cells(Index + PrviDO, "D").Value Like ("*rešetke*")) Then
            Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke & "_ST_VKL"
            Sheets("Sheet2").Cells(odmik, "B").Value = "DINT"
            Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & ", št. vkl."
            Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
            
            IndexR = IndexR + 2
            odmik = odmik + 1
    End If
    Index = Index + 1
Loop

'Kreiranje variabel za sinhronizacijo ure---------------------------------------------------------
message = "Predviden prvi register za sinhronizacijo ure. Po želji lahko spremeniš."
Title = "Predviden prvi register za sinhronizacijo ure."
sinhUreAddress = 3880
If sinhUreAddress <= IndexR Then
    sinhUreAddress = sinhUreAddress + 300
End If
Default = sinhUreAddress
sinhUreAddress = InputBox(message, Title, Default)

IndexR = sinhUreAddress
Index = 1

Do Until Index = 15
    Adresa = "%R" & Format(IndexR, "00000")
    Select Case Index
    Case 1
        Sheets("Sheet2").Cells(odmik, "A").Value = "PLC_COMM_CNT"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = "Števec za alarmiranje komunikacije, ne diraj!"
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
    Case 2
        Sheets("Sheet2").Cells(odmik, "A").Value = "URA_NA_PLC_LETO"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ""
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
    Case 3
        Sheets("Sheet2").Cells(odmik, "A").Value = "URA_NA_PLC_MESEC"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ""
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
    Case 4
        Sheets("Sheet2").Cells(odmik, "A").Value = "URA_NA_PLC_DAN"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ""
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
    Case 5
        Sheets("Sheet2").Cells(odmik, "A").Value = "URA_NA_PLC_URA"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ""
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
    Case 6
        Sheets("Sheet2").Cells(odmik, "A").Value = "URA_NA_PLC_MINUTA"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ""
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
    Case 7
        Sheets("Sheet2").Cells(odmik, "A").Value = "URA_NA_PLC_SEKUNDA"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ""
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
    Case 8
        Sheets("Sheet2").Cells(odmik, "A").Value = "URA_IZ_PLC_LETO"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ""
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
    Case 9
        Sheets("Sheet2").Cells(odmik, "A").Value = "URA_IZ_PLC_MESEC"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ""
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
    Case 10
        Sheets("Sheet2").Cells(odmik, "A").Value = "URA_IZ_PLC_DAN"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ""
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
    Case 11
        Sheets("Sheet2").Cells(odmik, "A").Value = "URA_IZ_PLC_URA"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ""
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
    Case 12
        Sheets("Sheet2").Cells(odmik, "A").Value = "URA_IZ_PLC_MINUTA"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ""
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
    Case 13
        Sheets("Sheet2").Cells(odmik, "A").Value = "URA_IZ_PLC_SEKUNDA"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ""
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
    Case 14
        Sheets("Sheet2").Cells(odmik, "A").Value = "URA_IZ_PLC_DAN_V_TEDNU"
        Sheets("Sheet2").Cells(odmik, "B").Value = "INT"
        Sheets("Sheet2").Cells(odmik, "C").Value = ""
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        IndexR = IndexR + 1
        odmik = odmik + 1
    Case Else
        IndexR = IndexR + 1
        odmik = odmik + 1
    End Select
    Index = Index + 1
    
Loop
'-------------------------------------------------------------------------------------------------

'KREIRANJE MEJ OBMOÈIJ ANALOGNIH VHODOV
'VNOS ZAÈETNEGA NASLOVA REGISTRA, KI SE UPORABLJA ZA TEMP
message = "Vnesi prvi register za uporabo temporary"
Title = "Vnos prvega registra za uporabo temporary"
Default = "10001"
temporary = InputBox(message, Title, Default)
Index = 1
IndexR = temporary
  y = DI("R", IndexR, PrviAI, SteviloAI, odmik, "_LC", ", LC")
  y = DI("R", IndexR, PrviAI, SteviloAI, odmik, "_UC", ", UC")
  y = DI("R", IndexR, PrviAI, SteviloAI, odmik, "_KOR", ", Korekcija")
  y = DI("R", IndexR, PrviAI, SteviloAI, odmik, "_WEIGHT", ", Utez")
 
 'KREIRANJE SPREMENLJIVKE REGULATORJA PID
 ''VNOS ZAÈETNEGA NASLOVA REGISTRA, KI SE UPORABLJA ZA PID REGULATORJE
message = "Predviden prvi register za uporabo PID regulatorja. Po želji lahko spremeniš."
Title = "Predviden prvi register za uporabo PID regulatorja"
PID = (Application.WorksheetFunction.RoundDown((IndexR / 1000), 0) * 1000) + 1
If PID <= IndexR Then
    PID = PID + 1000
End If
Default = PID
PID = InputBox(message, Title, Default)
Index = 1
IndexR = PID

Do Until Index = SteviloPID + 1
        
        Adresa = "%R" & Format(IndexR, "00000")
        ImeSistema = Sheets("IOT").Cells(Index + PrviPID, "B").Value
        ImeTocke = Sheets("IOT").Cells(Index + PrviPID, "C").Value
        opis = Sheets("IOT").Cells(Index + PrviPID, "D").Value
        
        Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_VA_" & ImeTocke
        Sheets("Sheet2").Cells(odmik, "B").Value = "WORD"
        Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis
        Sheets("Sheet2").Cells(odmik, "H").Value = 40
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
      
        Index = Index + 1
        IndexR = IndexR + 40
        odmik = odmik + 1
      
Loop

'KREIRANJE SPREMENLJIVK TIMERJEV ZA SISTEME
''VNOS ZAÈETNEGA NASLOVA REGISTRA, KI SE UPORABLJA ZA TIMERJE
message = "Predviden prvi register za uporabo timerjev. Po želji lahko spremeniš."
Title = "Predviden prvi register za uporabo timerjev"
Timer = (Application.WorksheetFunction.RoundDown((IndexR / 1000), 0) * 1000) + 13
If Timer <= IndexR Then
    Timer = Timer + 1000
End If
Default = Timer
Timer = InputBox(message, Title, Default)
Index = 1
IndexR = Timer

'ZAGONSKI TIMER
  y = DI("R", IndexR, PrviSis, SteviloSis, odmik, "_KVIT_TMR", "Timer kvitiranja")
 
 'KREIRANJE OSTALIH VREDNOSTI SIGNALOV (DIGITALNI SIGNALI)
Index = 1
konec = 0
Do Until konec = 1

    ImeTocke = Sheets("Sheet2").Cells(Index, "A").Value
    If ImeTocke = "" Then
        konec = 1
    Else
        konec = 0
        Index = Index + 1
    End If

    Test_di = Sheets("sheet2").Cells(Index, "B").Value Like "*BOOL*"

    If Test_di = True Then

        Sheets("Sheet2").Cells(Index, "E").Value = "YES"
        Sheets("Sheet2").Cells(Index, "F").Value = "NO"
        Sheets("Sheet2").Cells(Index, "G").Value = "On / Off"
        Sheets("Sheet2").Cells(Index, "H").Value = "0"
        Sheets("Sheet2").Cells(Index, "I").Value = "0"
        Sheets("Sheet2").Cells(Index, "J").Value = "NO"
        Sheets("Sheet2").Cells(Index, "K").Value = "NO"
        Sheets("Sheet2").Cells(Index, "L").Value = "1"
        Sheets("Sheet2").Cells(Index, "M").Value = "0"
        Sheets("Sheet2").Cells(Index, "N").Value = "GE FANUC PLC"

    End If

'KREIRANJE OSTALIH VREDNOSTI SIGNALOV (ANALOGNI SIGNALI)
    Test_di = Sheets("sheet2").Cells(Index, "B").Value Like "INT"
    If Test_di = True Then

        
        Sheets("Sheet2").Cells(Index, "E").Value = "YES"
        Sheets("Sheet2").Cells(Index, "F").Value = "NO"
        Sheets("Sheet2").Cells(Index, "G").Value = "Decimal"
        Sheets("Sheet2").Cells(Index, "H").Value = "0"
        Sheets("Sheet2").Cells(Index, "I").Value = "0"
        Sheets("Sheet2").Cells(Index, "J").Value = "NO"
        Sheets("Sheet2").Cells(Index, "K").Value = "NO"
        Sheets("Sheet2").Cells(Index, "L").Value = "16"
        Sheets("Sheet2").Cells(Index, "M").Value = "0"
        Sheets("Sheet2").Cells(Index, "N").Value = "GE FANUC PLC"

    End If

'KREIRANJE OSTALIH VREDNOSTI SIGNALOV (ANALOGNI SIGNALI)
    Test_di = Sheets("sheet2").Cells(Index, "B").Value Like "DINT"
    If Test_di = True Then

        Sheets("Sheet2").Cells(Index, "E").Value = "YES"
        Sheets("Sheet2").Cells(Index, "F").Value = "NO"
        Sheets("Sheet2").Cells(Index, "G").Value = "Decimal"
        Sheets("Sheet2").Cells(Index, "H").Value = "0"
        Sheets("Sheet2").Cells(Index, "I").Value = "0"
        Sheets("Sheet2").Cells(Index, "J").Value = "NO"
        Sheets("Sheet2").Cells(Index, "K").Value = "NO"
        Sheets("Sheet2").Cells(Index, "L").Value = "32"
        Sheets("Sheet2").Cells(Index, "M").Value = "0"
        Sheets("Sheet2").Cells(Index, "N").Value = "GE FANUC PLC"

    End If
    

'KREIRANJE OSTALIH VREDNOSTI SIGNALOV (PID REGULATORJI, TIMERJI)
    Test_di = Sheets("sheet2").Cells(Index, "B").Value Like "*WORD*"
    If Test_di = True Then

        Sheets("Sheet2").Cells(Index, "E").Value = "YES"
        Sheets("Sheet2").Cells(Index, "F").Value = "NO"
        Sheets("Sheet2").Cells(Index, "G").Value = "Decimal"
        Sheets("Sheet2").Cells(Index, "I").Value = "0"
        Sheets("Sheet2").Cells(Index, "J").Value = "NO"
        Sheets("Sheet2").Cells(Index, "K").Value = "NO"
        Sheets("Sheet2").Cells(Index, "L").Value = "16"
        Sheets("Sheet2").Cells(Index, "M").Value = "0"
        Sheets("Sheet2").Cells(Index, "N").Value = "GE FANUC PLC"

    End If

Loop
 
End Sub

'FUNKCIJA KREIRANJA DIGITALNIH TOÈK
' tip_tocke = npr: M, I,..
' index_tipa = zaporedna številka M ali I.....
' prvi_index_tipa = zaèetni naslov I ali M ali Q...
' stevilo_tock = število toèk doloèenega tipa npr: 24 DI,...
Function DI(tip_tocke, index_tipa, prvi_index_tipa, stevilo_tock, odmik, koncnica, koncnica_opisa)
 Dim Index As Integer
 
 Index = 1
  If tip_tocke = "M" Then
        tip_tocke_1 = "VD"
        tip_tocke_2 = "BOOL"
        End If
        If tip_tocke = "I" Then
        tip_tocke_1 = "DI"
        tip_tocke_2 = "BOOL"
        End If
        If tip_tocke = "AI" Then
        tip_tocke_1 = "AI"
        tip_tocke_2 = "INT"
        End If
        If tip_tocke = "AQ" Then
        tip_tocke_1 = "AO"
        tip_tocke_2 = "INT"
        End If
        If tip_tocke = "R" Then
        tip_tocke_1 = "VA"
        tip_tocke_2 = "INT"
        End If
        If tip_tocke = "Q" Then
        tip_tocke_1 = "DO"
        tip_tocke_2 = "BOOL"
        End If
        If tip_tocke = "VDQ" Then
        tip_tocke = "Q"
        tip_tocke_1 = "VD"
        tip_tocke_2 = "BOOL"
        End If
        If tip_tocke = "RD" Then
        tip_tocke_1 = "VA"
        tip_tocke_2 = "DINT"
        End If
   Do Until Index = stevilo_tock + 1
   
        If tip_tocke Like "RD" Then
            Adresa = "%R" & Format(index_tipa, "00000")
        ElseIf tip_tocke Like "AI" Or tip_tocke Like "AQ" Then
            Adresa = "%" & tip_tocke & Format(index_tipa, "0000")
        Else
            Adresa = "%" & tip_tocke & Format(index_tipa, "00000")
        End If
        
        ImeSistema = Sheets("IOT").Cells(Index + prvi_index_tipa, "B").Value
        ImeTocke = Sheets("IOT").Cells(Index + prvi_index_tipa, "C").Value
        opis = Sheets("IOT").Cells(Index + prvi_index_tipa, "D").Value
               
        Test_di = Sheets("IOT").Cells(Index + prvi_index_tipa, "a").Value Like "*%SIS*"
        If Test_di = True Then
 '          Sheets("Sheet2").Cells(ODMIK, "A").Value = ImeSistema & "_" & tip_tocke_1 & koncnica krt
           Sheets("Sheet2").Cells(odmik, "A").Value = ImeTocke & "_" & tip_tocke_1 & koncnica
 '          Sheets("Sheet2").Cells(ODMIK, "C").Value = ImeSistema & ": " & koncnica_opisa   krt
           Sheets("Sheet2").Cells(odmik, "C").Value = ImeTocke & ": " & koncnica_opisa
           Else
              If Sheets("IOT").Cells(Index + prvi_index_tipa, "C").Value = "" Then
              
              Sheets("Sheet2").Cells(odmik, "A").Value = tip_tocke & index_tipa
              Else
              Sheets("Sheet2").Cells(odmik, "A").Value = ImeSistema & "_" & tip_tocke_1 & "_" & ImeTocke & koncnica
              End If
        
                
              If Sheets("IOT").Cells(Index + prvi_index_tipa, "D").Value = "" Then
                Sheets("Sheet2").Cells(odmik, "C").Value = ""
              Else
                Sheets("Sheet2").Cells(odmik, "C").Value = ImeSistema & ": " & opis & koncnica_opisa
              End If
        
        End If
        
        Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
        
      Test_di = Sheets("Sheet2").Cells(odmik, "A").Value Like "*_TMR"
      If Test_di = True Then
      Sheets("Sheet2").Cells(odmik, "B").Value = "WORD"
      Sheets("Sheet2").Cells(odmik, "H").Value = 3
      index_tipa = index_tipa + 3
      Else
      Sheets("Sheet2").Cells(odmik, "B").Value = tip_tocke_2
        If tip_tocke_2 Like "DINT" Then
            index_tipa = index_tipa + 2
        Else
            index_tipa = index_tipa + 1
        End If
    End If
      
      Index = Index + 1
      odmik = odmik + 1
     
Loop
     
 End Function
 
 
 
 Function E(index_tipa, krmilnik, odmik, koncnica, koncnica_opisa)
 
    Adresa = "%M" & Format(index_tipa, "00000")
    Sheets("Sheet2").Cells(odmik, "A").Value = krmilnik & "_VD_" & koncnica
    Sheets("Sheet2").Cells(odmik, "B").Value = "BOOL"
    Sheets("Sheet2").Cells(odmik, "C").Value = krmilnik & ": " & koncnica_opisa
    Sheets("Sheet2").Cells(odmik, "O").Value = Adresa
    
    IndexM = IndexM + 1
    odmik = odmik + 1
    index_tipa = index_tipa + 1
    
 End Function

