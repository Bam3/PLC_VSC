Attribute VB_Name = "MyRibbon"

'namespace=vba-files/ribbons

Public Sub Variables(ByRef control As Office.IRibbonControl)
    Call ustvari_csv
End Sub

Public Sub AI2EGU(ByRef control As Office.IRibbonControl)
    Call AI2EGU_Create
End Sub

Public Sub MEJE(ByRef control As Office.IRibbonControl)
    Call MEJE_Create
End Sub

Public Sub EGUALM(ByRef control As Office.IRibbonControl)
    Call EGUALM_Create
End Sub

Public Sub CTRL(ByRef control As Office.IRibbonControl)
    Call CTRL_Create()
End Sub

Public Sub I2VD(ByRef control As Office.IRibbonControl)
    Call I2VD_Create()
End Sub