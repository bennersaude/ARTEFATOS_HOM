'HASH: 0B7609AF213110284C055D0ED63FCEED
 
 
Public Sub IMPORTAR_OnClick() 
  Dim obj As Object 
  Set obj = CreateBennerObject("CS.ReportFunctions") 
 
  obj.ImportDetail(CurrentSystem, RecordHandleOfTable("R_RELATORIOS")) 
 
  Set ob = Nothing 
End Sub 
