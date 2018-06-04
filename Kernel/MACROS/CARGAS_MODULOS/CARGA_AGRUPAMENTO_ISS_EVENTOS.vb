'HASH: A09BF0402372E1576598E59B3DB3A04B
Public Sub BOTAOGERAREVENTO_OnClick()

  Dim Obj As Object
  Set Obj = CreateBennerObject("BSANS001.uRotinas")
  Obj.GerarEventoClasse(CurrentSystem, RecordHandleOfTable("SFN_AGRUPAMENTOISS"), 1)
  Set Obj = Nothing

End Sub
