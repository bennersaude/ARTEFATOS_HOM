'HASH: 27F30166AC6DE40C5E465462CB16AE81

Public Sub Main
  pPeg = CInt(ServiceVar("pPeg"))

  Dim callEntity As CSEntityCall

  Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.ProcessamentoContas.SamPeg, Benner.Saude.Entidades", "VerificarBloqueioAlteracoesPegGuiaEvento")
  callEntity.AddParameter(pdtInteger, pPeg)

  ServiceResult = CBool(callEntity.Execute)

  Set callEntity = Nothing
End Sub
