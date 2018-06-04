'HASH: 3EB4031EECC1706B0EF059C2935932FA

Public Sub Main
  Dim GeoLocalizacao As CSBusinessComponent
  Set GeoLocalizacao = BusinessComponent.CreateInstance("Benner.Saude.Diversos.Localizacao.GeoLocalizacaoBll, Benner.Saude.Diversos.Localizacao")

  GeoLocalizacao.AddParameter(pdtString, CStr(SessionVar("ATUALIZARTUDO")))
  GeoLocalizacao.AddParameter(pdtString, CStr(SessionVar("TIPOATUALIZACAO")))
  GeoLocalizacao.Execute("AtualizarPeloAgendamento")

  Set GeoLocalizacao = Nothing
End Sub
