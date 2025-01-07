VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OsnovnaFormaESIGN 
   Caption         =   "Vpis zaèetnih parametrov izdelave Esign"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17010
   OleObjectBlob   =   "OsnovnaFormaESIGN.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OsnovnaFormaESIGN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CheckBox_ALL_Click()
    If CheckBox_ALL = True Then
        CheckBox_A_EN.Value = True
        CheckBox_KVIT.Value = True
        CheckBox_ZAK1.Value = True
        CheckBox_ZAK2.Value = True
        CheckBox_HIHI.Value = True
        CheckBox_HI.Value = True
        CheckBox_LOLO.Value = True
        CheckBox_LO.Value = True
        CheckBox_PID_ROCNO.Value = True
        CheckBox_RAMP.Value = True
        CheckBox_VA_PID.Value = True
        CheckBox_AO.Value = True
        CheckBox_KVIT_SCADA = True
        CheckBox_VKLOP_SCADA = True
        CheckBox_AUTO = True
        CheckBox_ROCNO = True
        CheckBox_SERVIS = True
        CheckBox_OBRH_ST_VKL = True
        CheckBox_DI_SRV_SB = True
        CheckBox_DI_SRV_SV = True
        CheckBox_REZIM_ACT = True
        CheckBox_DI_MN = True
        CheckBox_DI_SR = True
        CheckBox_VA_MN = True
        CheckBox_VA_SR = True
    Else
        CheckBox_A_EN.Value = False
        CheckBox_KVIT.Value = False
        CheckBox_ZAK1.Value = False
        CheckBox_ZAK2.Value = False
        CheckBox_HIHI.Value = False
        CheckBox_HI.Value = False
        CheckBox_LOLO.Value = False
        CheckBox_LO.Value = False
        CheckBox_PID_ROCNO.Value = False
        CheckBox_RAMP.Value = False
        CheckBox_VA_PID.Value = False
        CheckBox_AO.Value = False
        CheckBox_KVIT_SCADA = False
        CheckBox_VKLOP_SCADA = False
        CheckBox_AUTO = False
        CheckBox_ROCNO = False
        CheckBox_SERVIS = False
        CheckBox_OBRH_ST_VKL = False
        CheckBox_DI_SRV_SB = False
        CheckBox_DI_SRV_SV = False
        CheckBox_REZIM_ACT = False
        CheckBox_DI_MN = False
        CheckBox_DI_SR = False
        CheckBox_VA_MN = False
        CheckBox_VA_SR = False
    End If
End Sub
Public Sub CommandButton1_Click()
Dim PLCName, Objekt, Location As String
Dim AREA As Variant
Dim AREA1, AREA2, AREA3, AREA4, AREA5, AREA6, AREA7, AREA8, AREA9, AREA10 As String
Dim sklop As String
'global Counter2 As Integer
 
'zbrišemo in ustvarimo Sheet
Application.DisplayAlerts = False
On Error Resume Next
Sheets("ESIGN").Delete
Sheets("ESIGN_TAB").Delete
Sheets.Add.Name = "ESIGN"
Sheets.Add.Name = "ESIGN_TAB"

Counter2 = 1
PLCName = TextBox_PLC.Value
Objekt = TextBox_OBJEKT.Value
Location = TextBox_Location.Value
sklop = TextBox_sistem.Value

AREA1 = Location
AREA2 = Objekt
'AREA3 = "RMS"
AREA4 = sklop
AREA5 = "N.A."
AREA7 = PLCName
AREA8 = "N.A."
AREA9 = "N.A."
AREA10 = "N.A."

AREA = Array(AREA1, AREA2, AREA3, AREA4, AREA5, AREA6, AREA7, AREA8, AREA9, AREA10)

If CheckBox_A_EN Then Counter2 = A_EN(AREA, Counter2, sklop, ComboBoxSheet.Value, "ALM_PAR")
If CheckBox_KVIT Then Counter2 = KVIT(AREA, Counter2, sklop, ComboBoxSheet.Value, "ALM_PAR")

If CheckBox_HIHI Then Counter2 = ALM(AREA, Counter2, sklop, ComboBoxSheet.Value, "HIHI", "ALM_PAR")
If CheckBox_HI Then Counter2 = ALM(AREA, Counter2, sklop, ComboBoxSheet.Value, "HI", "ALM_PAR")
If CheckBox_LO Then Counter2 = ALM(AREA, Counter2, sklop, ComboBoxSheet.Value, "LO", "ALM_PAR")
If CheckBox_LOLO Then Counter2 = ALM(AREA, Counter2, sklop, ComboBoxSheet.Value, "LOLO", "ALM_PAR")

If CheckBox_ZAK1 Then Counter2 = ZAK_AI(AREA, Counter2, sklop, ComboBoxSheet.Value, "ZAK1", "ALM_PAR")
If CheckBox_ZAK2 Then Counter2 = ZAK_AI(AREA, Counter2, sklop, ComboBoxSheet.Value, "ZAK2", "ALM_PAR")

If CheckBox_PID_ROCNO Then Counter2 = PID_ROCNO(AREA, Counter2, sklop, ComboBoxSheet.Value, "REG_PAR")
If CheckBox_RAMP Then Counter2 = RAMP(AREA, Counter2, sklop, ComboBoxSheet.Value, "REG_PAR")
If CheckBox_VA_PID Then Counter2 = VA_PID(AREA, Counter2, sklop, ComboBoxSheet.Value, "REG_PAR")
If CheckBox_AO Then Counter2 = AO(AREA, Counter2, sklop, ComboBoxSheet.Value, "REG_PAR")

If CheckBox_KVIT_SCADA Then Counter2 = SCADA_ESIGN(AREA, Counter2, sklop, ComboBoxSheet.Value, "KVIT_SCADA", "SIS_PAR")
If CheckBox_VKLOP_SCADA Then Counter2 = SCADA_ESIGN(AREA, Counter2, sklop, ComboBoxSheet.Value, "VKLOP_SCADA", "SIS_PAR")
If CheckBox_AUTO Then Counter2 = SCADA_ESIGN(AREA, Counter2, sklop, ComboBoxSheet.Value, "AUTO", "SIS_PAR")
If CheckBox_ROCNO Then Counter2 = SCADA_ESIGN(AREA, Counter2, sklop, ComboBoxSheet.Value, "ROCNO", "SIS_PAR")
If CheckBox_SERVIS Then Counter2 = SCADA_ESIGN(AREA, Counter2, sklop, ComboBoxSheet.Value, "SERVIS", "SIS_PAR")
If CheckBox_OBRH_ST_VKL Then Counter2 = OBRH_ST_VKL(AREA, Counter2, sklop, ComboBoxSheet.Value, "SIS_PAR")

If CheckBox_DI_SRV_SB Then Counter2 = DI_SRV(AREA, Counter2, sklop, ComboBoxSheet.Value, "SB", "SIS_PAR")
If CheckBox_DI_SRV_SV Then Counter2 = DI_SRV(AREA, Counter2, sklop, ComboBoxSheet.Value, "SV", "SIS_PAR")

If CheckBox_REZIM_ACT Then Counter2 = REZ_ACT(AREA, Counter2, sklop, ComboBoxSheet.Value, "SIS_PAR")

If CheckBox_DI_MN Then Counter2 = DI_MAN_SRV(AREA, Counter2, sklop, ComboBoxSheet.Value, "MN", "SIS_PAR")
If CheckBox_DI_SR Then Counter2 = DI_MAN_SRV(AREA, Counter2, sklop, ComboBoxSheet.Value, "SR", "SIS_PAR")
If CheckBox_VA_MN Then Counter2 = VA_MAN_SRV(AREA, Counter2, sklop, ComboBoxSheet.Value, "MN", "SIS_PAR")
If CheckBox_VA_MN Then Counter2 = VA_MAN_SRV(AREA, Counter2, sklop, ComboBoxSheet.Value, "SR", "SIS_PAR")

Sheets(2).UsedRange.Columns.AutoFit
Unload Me

End Sub



Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ComboBoxSheet.AddItem ws.Name
    Next ws
    'dej prvega v prikaz
    ComboBoxSheet.Value = "TGD"
    CheckBox_ALL = True
    TextBox_Location.Value = Sheets("ESIGN_SETTINGS").Cells(6, "A").Value
    TextBox_OBJEKT.Value = Sheets("ESIGN_SETTINGS").Cells(6, "B").Value
    TextBox_sistem.Value = Sheets("ESIGN_SETTINGS").Cells(6, "D").Value
    TextBox_PLC.Value = Sheets("ESIGN_SETTINGS").Cells(6, "G").Value
    
End Sub
