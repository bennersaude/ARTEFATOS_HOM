'HASH: 7253B63D59E9346EF56BD6AF164A3F09
'#Uses "*bsShowMessage"


Public Sub TABELABRASINDICE_OnPopup(ShowPopup As Boolean)
  Dim voDll As Object
  Dim handlexx As Long
  ShowPopup = False
  Set voDll = CreateBennerObject("Procura.Procurar")
  handlexx = voDll.Exec(CurrentSystem, "TIS_TABELAPRECO", "CODIGO|DESCRICAO", 2, "Código|Descrição", "VERSAOTISS = " + CurrentQuery.FieldByName("HANDLE").AsString, "Tabelas de preço", True, "")
  If handlexx <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("TABELABRASINDICE").Value = handlexx
  End If
  Set voDll = Nothing
End Sub

Public Sub TABELASIMPRO_OnPopup(ShowPopup As Boolean)
  Dim voDll As Object
  Dim handlexx As Long
  ShowPopup = False
  Set voDll = CreateBennerObject("Procura.Procurar")
  handlexx = voDll.Exec(CurrentSystem, "TIS_TABELAPRECO", "CODIGO|DESCRICAO", 2, "Código|Descrição", "VERSAOTISS = " + CurrentQuery.FieldByName("HANDLE").AsString, "Tabelas de preço", True, "")
  If handlexx <> 0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("TABELASIMPRO").Value = handlexx
  End If
  Set voDll = Nothing
End Sub
