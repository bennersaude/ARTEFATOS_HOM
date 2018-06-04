'HASH: 39DB9D51FF617612F1A2315AEE827753

Public Sub BOTAOGERAR_OnClick()
  Dim Obj As Object


  Set Obj = CreateBennerObject("SamGerarClasseGerencial.Rotinas")
  Obj.Gerar(CurrentSystem, CurrentQuery.FieldByName("ITEM").AsInteger)
  Set Obj = Nothing

End Sub


Public Sub BOTAOEXCLUIR_OnClick()
  Dim Obj As Object


  Set Obj = CreateBennerObject("SamGerarClasseGerencial.Rotinas")
  Obj.Excluir(CurrentSystem, CurrentQuery.FieldByName("ITEM").AsInteger)
  Set Obj = Nothing

End Sub

