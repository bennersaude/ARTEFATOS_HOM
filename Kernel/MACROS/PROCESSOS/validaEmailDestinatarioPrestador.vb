'HASH: CBAA7D7DF32566EF7120037AD2DC856E

Public Sub Main

	Dim ValidarEmail As CSBusinessComponent

	Set ValidarEmail = BusinessComponent.CreateInstance("Benner.Saude.ProcessamentoContas.Business.TabelaVirtual.TvForm0148BLL, Benner.Saude.ProcessamentoContas.Business")
	ValidarEmail.AddParameter(pdtString, ServiceVar("pEmail"))

	ServiceResult = ValidarEmail.Execute("ValidaEmailDestinatarioPrestador")

	Set ValidarEmail = Nothing

End Sub
