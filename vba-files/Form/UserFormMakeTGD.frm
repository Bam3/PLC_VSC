VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMakeTGD 
   Caption         =   "Ustvari TGD"
   ClientHeight    =   10170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4920
   OleObjectBlob   =   "UserFormMakeTGD.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormMakeTGD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub createButton_Click()
    Call exportPLC_to_TDG(NodeNameTxtBox.Value, _
                          PLCnameTxtBox.Value, _
                          registerRTxtBox.Value, _
                          registerMTxtBox.Value, _
                          registerQTxtBox.Value, _
                          Predpona.Value, _
                          zacAO.Value, _
                          koncAO.Value, _
                          suffixUsageChk.Value, _
                          SecurityUsageChk.Value, _
                          OptionButtonAR0.Value, _
                          OptionButtonDRQ1.Value, _
                          ComboBoxSheet.Value)
End Sub
Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ComboBoxSheet.AddItem ws.Name
    Next ws
    'dej prvega v prikaz
    ComboBoxSheet.Value = "Sheet2"

    OptionButtonAR0.Value = True
    OptionButtonDRQ1.Value = True
End Sub

