'HASH: 726A73C86AA308F6DC56E1D0A722C17F
 

Public Sub TABLE_AfterScroll()
 Dim Retorno As String
 Dim component As CSBusinessComponent

 Set component = BusinessComponent.CreateInstance("Benner.Saude.Beneficiarios.UtilizacaoDeServicos.Business.PeriodosDeUtilizacaoBusiness, Benner.Saude.Beneficiarios.UtilizacaoDeServicos.Business")
 Retorno = component.Execute("ConsultaPeriodos")

 CurrentQuery.FieldByName("RESULTADO").AsString = Retorno
 Set componente = Nothing

End Sub
