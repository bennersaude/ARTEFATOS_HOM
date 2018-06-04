'HASH: 36A449CF29AE13B8906D9427F5ED1155
'#Uses "*ProcuraEvento"
'#Uses "*ProcuraTabelaUS"
'#Uses "*ProcuraTabelaFilme"
'#Uses "*bsShowMessage"

Option Explicit

Public Sub Incluir_OnClick()
	Dim Interface As Object
	Dim psMensagem As String
	Dim handlePrestadores As String

	Set Interface =CreateBennerObject("BSINTERFACE0001.BUSCAPRESTADOR")

	Interface.Abrir(CurrentSystem, psMensagem, 0, "", "", 0, 0, "", True, handlePrestadores)

	Set Interface = Nothing

	If handlePrestadores <> "" Then

		Dim business As CSBusinessComponent
		Dim resposta As String

		Set business = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.Excepcionalidade.SamExcepcionalidadePrestadorBLL, Benner.Saude.Prestadores.Business")
		business.AddParameter(pdtInteger, RecordHandleOfTable("SAM_EXCEPCIONALIDADE"))
		business.AddParameter(pdtString, handlePrestadores)
		resposta = business.Execute("IncluirPrestadores")

		MsgBox(resposta)

		Set business = Nothing

		RefreshNodesWithTable "SAM_EXCEPCIONALIDADE_PRESTADOR"
	End If
End Sub
