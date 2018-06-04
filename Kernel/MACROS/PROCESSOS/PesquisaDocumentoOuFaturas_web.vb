'HASH: EFF9C7FA9779EC7CD004E40870C862D6
Public Sub Main
	' Codifique aqui o método principal
Dim acao As Long
acao = CLng(ServiceVar("acao"))  '1 para faturas e 2 para documentos

Dim handleContaFinanceira,  dataI,  dataF,  rdEmissao,  rdVencimento,  chkBaixa,  chkCancel,  chkAberta As String
Dim interface As CSBusinessComponent

Select Case acao
  Case 1  'faturas
     Set interface = BusinessComponent.CreateInstance("Benner.Saude.Financeiro.Business.SfnFaturaBLL, Benner.Saude.Financeiro.Business")

     handleContaFinanceira =CStr(ServiceVar("handleContaFinanceira"))
     dataI                 =CStr(ServiceVar("dataI"))
     dataF                 =CStr(ServiceVar("dataF"))
     rdEmissao             =CStr(ServiceVar("rdEmissao"))
     rdVencimento          =CStr(ServiceVar("rdVencimento"))
     chkBaixa              =CStr(ServiceVar("chkBaixa"))
     chkCancel             =CStr(ServiceVar("chkCancel"))
     chkAberta             =CStr(ServiceVar("chkAberta"))

     interface.AddParameter(pdtString, handleContaFinanceira)
	 interface.AddParameter(pdtString, dataI)
	 interface.AddParameter(pdtString, dataF)
	 interface.AddParameter(pdtString, rdEmissao)
	 interface.AddParameter(pdtString, rdVencimento)
	 interface.AddParameter(pdtString, chkBaixa)
	 interface.AddParameter(pdtString, chkCancel)
	 interface.AddParameter(pdtString, chkAberta)

     ServiceResult = interface.Execute("SelecionarFaturasTelaCustomizadaXml")
     Set interface = Nothing


 Case 2  'documentos
     Set interface = BusinessComponent.CreateInstance("Benner.Saude.Financeiro.Business.SfnDocumentoBLL, Benner.Saude.Financeiro.Business")

     handleContaFinanceira =CStr(ServiceVar("handleContaFinanceira"))
     dataI                 =CStr(ServiceVar("dataI"))
     dataF                 =CStr(ServiceVar("dataF"))
     rdEmissao             =CStr(ServiceVar("rdEmissao"))
     rdVencimento          =CStr(ServiceVar("rdVencimento"))
     chkBaixa              =CStr(ServiceVar("chkBaixa"))
     chkCancel             =CStr(ServiceVar("chkCancel"))
     chkAberta             =CStr(ServiceVar("chkAberta"))

     interface.AddParameter(pdtString, handleContaFinanceira)
	 interface.AddParameter(pdtString, dataI)
	 interface.AddParameter(pdtString, dataF)
	 interface.AddParameter(pdtString, rdEmissao)
	 interface.AddParameter(pdtString, rdVencimento)
	 interface.AddParameter(pdtString, chkBaixa)
	 interface.AddParameter(pdtString, chkCancel)
	 interface.AddParameter(pdtString, chkAberta)

     ServiceResult = interface.Execute("SelecionarDocumentosTelaCustomizadaXml")
     Set interface = Nothing

End Select

End Sub
