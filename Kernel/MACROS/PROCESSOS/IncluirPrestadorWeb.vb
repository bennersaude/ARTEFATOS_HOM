'HASH: 3F903420040A5D7088C33580C3981E5E
Public Sub Main
	' Codifique aqui o método principal
	Dim acao As Long
	acao = CLng(ServiceVar("acao"))
	
'selecionarPrestadoresNaRedeAtendimentoPortalXml(string handleRedeAtendimento, string pesquisa, string campo, string tipo)
Dim interface As CSBusinessComponent
Set interface = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.RedeAtendimentoPortal.SamRedeatendportalPrestadorBLL, Benner.Saude.Prestadores.Business")

Select Case acao
  Case 1
     Dim handleRedeAtendimento,pesquisa,campo,tipo As String

     handleRedeAtendimento =CStr(ServiceVar("handleRedeAtendimento"))
     pesquisa              =CStr(ServiceVar("pesquisa"))
     campo                 =CStr(ServiceVar("campo"))
     tipo                  =CStr(ServiceVar("tipo"))

     interface.AddParameter(pdtString, handleRedeAtendimento)
     interface.AddParameter(pdtString, pesquisa)
     interface.AddParameter(pdtString, campo)
     interface.AddParameter(pdtString, tipo)
     ServiceResult = interface.Execute("SelecionarPrestadoresNaRedeAtendimentoPortalXml")

 Case 2
     Dim handleRedeAtendimentoI,handlesI As String

     handleRedeAtendimentoI =CStr(ServiceVar("handleRedeAtendimento"))
     handlesI               =CStr(ServiceVar("handles"))

     interface.AddParameter(pdtString, handleRedeAtendimentoI)
     interface.AddParameter(pdtString, handlesI)
     ServiceResult = interface.Execute("IncluirPrestadoresNaRedeAtendimentoPortal")

End Select

Set interface = Nothing
End Sub
