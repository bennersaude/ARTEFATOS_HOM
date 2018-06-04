'HASH: 6AF655234D2AC515D5CCD2A76663D7EF

Public Sub Main

	Dim TraduzirMensagem As CSBusinessComponent

	Set TraduzirMensagem = BusinessComponent.CreateInstance("Benner.Saude.ProcessamentoContas.Business.TabelaVirtual.TvForm0148BLL, Benner.Saude.ProcessamentoContas.Business")
	TraduzirMensagem.AddParameter(pdtString, "P")
	TraduzirMensagem.AddParameter(pdtInteger, CInt(ServiceVar("pPeg")))
	TraduzirMensagem.AddParameter(pdtInteger, CInt(ServiceVar("pMotivoGlosa")))
	TraduzirMensagem.AddParameter(pdtString, ServiceVar("pMensagemHtml"))

	ServiceResult = TraduzirMensagem.Execute("TraduzirMensagemPadrao")

	Set TraduzirMensagem = Nothing

End Sub
