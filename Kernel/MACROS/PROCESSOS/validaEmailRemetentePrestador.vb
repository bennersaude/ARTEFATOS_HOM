'HASH: 59714A54B880836C2B1B63B54724F80C

Public Sub Main

	Dim ValidarEmail As CSBusinessComponent

	Set ValidarEmail = BusinessComponent.CreateInstance("Benner.Saude.ProcessamentoContas.Business.TabelaVirtual.TvForm0148BLL, Benner.Saude.ProcessamentoContas.Business")
	ValidarEmail.AddParameter(pdtString, ServiceVar("pEmail"))

	ServiceResult = ValidarEmail.Execute("ValidaEmailRemetentePrestador")

	Set ValidarEmail = Nothing

End Sub
