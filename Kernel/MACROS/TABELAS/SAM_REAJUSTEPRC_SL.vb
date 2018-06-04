'HASH: 5D52D84BE36A3BC69760575545B3E1B3

Option Explicit
'#Uses "*ProcuraTabelaUS"
'#Uses "*bsShowMessage"


Public Sub TABELAUS_OnPopup(ShowPopup As Boolean)

  Dim vHandle As Long
  ShowPopup = False
  vHandle = ProcuraTabelaUS(TABELAUS.Text)
  If vHandle<>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("TABELAUS").Value = vHandle
  End If

End Sub

Public Sub TABLE_AfterPost()
  RefreshNodesWithTable("SAM_REAJUSTEPRC_SL")
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vAssocociacaodaTabela As Long

  vAssocociacaodaTabela = CurrentQuery.State
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > _
                              CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
    bsShowMessage("Data INICIAL não pode ser maior que a data FINAL", "E")
    CanContinue = False
  ElseIf CurrentQuery.FieldByName("NOVAVIGENCIA").AsDateTime <= _
                                    CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
    bsShowMessage("NOVA VIGÊNCIA deve ser maior que a data FINAL", "E")
    CanContinue = False
  End If
End Sub


