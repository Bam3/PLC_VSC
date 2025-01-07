VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSumniki 
   Caption         =   "Izberi sheet za popravljanje šumnikov"
   ClientHeight    =   1770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4350
   OleObjectBlob   =   "UserFormSumniki.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormSumniki"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
    Dim wsName As String
    wsName = ComboBox1.Value
    If wsName <> "" Then
        popraviSumnike ThisWorkbook.Worksheets(wsName)
        Unload Me
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ComboBox1.AddItem ws.Name
    Next ws
    'dej prvega v prikaz
    ComboBox1.ListIndex = 0
End Sub
