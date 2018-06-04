'HASH: D7E470818C2FF80EF7A6BAA12EAEC888
 Public Sub BOTAOGERAR_OnClick()
  Dim Obj As Object
  Set Obj = CreateBennerObject("SamGerarEventos.Rotinas")
  Obj.Gerar(CurrentSystem, RecordHandleOfTable("GER_ITEM"))
  Set Obj = Nothing
End Sub
Public Sub BOTAOEXCLUIR_OnClick()
  Dim Obj As Object
  Set Obj = CreateBennerObject("SamGerarEventos.Rotinas")
  Obj.Excluir(CurrentSystem, RecordHandleOfTable("GER_ITEM"))
  Set Obj = Nothing
End Sub
