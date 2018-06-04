'HASH: 38865C7E350ED6EF416AF1484DCE17C4

Public Sub BOTAOGERAR_OnClick()
  Dim Obj As Object


  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Gerar(CurrentSystem, "GER_GRUPOCCUSTO_CCUSTO", "Duplicando Centro de Custo para o Grupo", "SFN_CENTROCUSTO", "CCUSTO", "GRUPOCCUSTO", CurrentQuery.FieldByName("GRUPOCCUSTO").AsInteger, "S", "ESTRUTURA")
  Set Obj = Nothing

End Sub


Public Sub BOTAOEXCLUIR_OnClick()
  Dim Obj As Object


  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Excluir(CurrentSystem, "GER_GRUPOCCUSTO_CCUSTO", "Excluindo Centro de Custo para o Grupo", "SFN_CENTROCUSTO", "CCUSTO", "GRUPOCCUSTO", CurrentQuery.FieldByName("GRUPOCCUSTO").AsInteger, "S", "ESTRUTURA")
  Set Obj = Nothing

End Sub

