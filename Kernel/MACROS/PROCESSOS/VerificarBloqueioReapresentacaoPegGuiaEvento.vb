'HASH: EE9CA4E9F6589EDBC7F4A7488749E7C8

Public Sub Main

  pPeg = CInt(ServiceVar("pPeg"))

  Dim callEntity As CSEntityCall

  Set callEntity = BusinessEntity.CreateCall("Benner.Saude.Entidades.ProcessamentoContas.SamPeg, Benner.Saude.Entidades", "VerificarBloqueioReapresentacaoPegGuiaEvento")
  callEntity.AddParameter(pdtInteger, pPeg)

  ServiceResult = CBool(callEntity.Execute)

  Set callEntity = Nothing
End Sub
