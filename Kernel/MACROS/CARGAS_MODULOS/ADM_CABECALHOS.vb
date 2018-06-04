'HASH: DFCD2A00237EB8301A29FD39F6E224BC
 
 
Public Sub IMPORTAR_OnClick() 
  Dim obj As Object 
  Set obj = CreateBennerObject("CS.ReportFunctions") 
 
  obj.ImportDetail(CurrentSystem, RecordHandleOfTable("R_RELATORIOS")) 
 
  Set obj = Nothing 
End Sub 
