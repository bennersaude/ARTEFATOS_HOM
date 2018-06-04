'HASH: 3E295095299A91EFC29BE839F7DE7FAA
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterPost()

  Dim componente As CSBusinessComponent

  Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.NegociacaoPreco.TvCalcularIndiceAcumuladoBLL, Benner.Saude.Prestadores.Business")
  componente.AddParameter(pdtDateTime, CurrentQuery.FieldByName("DATAINICIAL").AsDateTime)
  componente.AddParameter(pdtDateTime, CurrentQuery.FieldByName("DATAFINAL").AsDateTime)
  bsshowmessage(componente.Execute("CalcularIndicesAcumulados"), "I")

  Set componente = Nothing
End Sub

