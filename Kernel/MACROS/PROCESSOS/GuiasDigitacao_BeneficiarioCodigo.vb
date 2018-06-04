'HASH: 038DF8731528C5A165A4F0A00B88BC79

'# Uses "*addXMLAtt"
Public Sub Main


	Dim psDigitado As String
	Dim handle As Long

	psDigitado = CStr( ServiceVar("psDigitado") )

	Dim sp As BStoredProc
	Set sp=NewStoredProc
	Dim sql As BPesquisa
	Dim xml As String

	On Error GoTo erro

	sp.Name="BSBen_BuscaBenefPrioridade"
	sp.AddParam("pv_ValorBuscar",ptInput, ftString, 100)
	sp.AddParam("pi_HandleBeneficiario",ptOutput, ftInteger)
	sp.ParamByName("pv_ValorBuscar").AsString = psDigitado

	sp.ExecProc
	handle = sp.ParamByName("pi_HandleBeneficiario").AsInteger

	If handle > 0 Then
		Set sql=NewQuery
		sql.Add("SELECT B.HANDLE, B.BENEFICIARIO, B.NOME, P.DESCRICAO PLANO, M.CARTAONACIONALSAUDE CNS, CART.DATAFINALVALIDADE VALIDADECARTAO, M.SEXO,B.DATACANCELAMENTO")
		sql.Add("FROM SAM_BENEFICIARIO B")
		sql.Add("JOIN SAM_CONTRATO C ON (C.HANDLE=B.CONTRATO)")
		sql.Add("JOIN SAM_PLANO P ON (P.HANDLE=C.PLANO)")
		sql.Add("JOIN SAM_MATRICULA M ON (M.HANDLE=B.MATRICULA)")
		sql.Add("LEFT JOIN SAM_BENEFICIARIO_CARTAOIDENTIF CART ON (CART.BENEFICIARIO=B.HANDLE AND CART.SITUACAO='N')")
		sql.Add("WHERE B.HANDLE=:HANDLE ")
		sql.ParamByName("HANDLE").AsInteger=handle
		sql.Active=True

		xml="<registros>"
		While Not sql.EOF
			xml=xml + "<registro>"
			xml=xml + addXMLAtt( "handle", "handle", sql, "")
			xml=xml + addXMLAtt( "nome", "nome", sql, "caption='Nome' width='300'")
			xml=xml + addXMLAtt( "codigo", "beneficiario", sql, "caption='Código' width='120'")
			xml=xml + addXMLAtt( "plano", "plano", sql, "caption='Plano' width='200'")
			xml=xml + addXMLAtt( "cns", "cns", sql, "caption='CNS' width='90'")
			xml=xml + addXMLAtt( "validadeCartao", "validadeCartao", sql, "caption='Validade cartão' width='90'")
			xml=xml + addXMLAtt( "dataCancelamento", "dataCancelamento", sql, "caption='Data de Cancelamento' width='110'")
			xml=xml + addXMLAtt( "sexo", "sexo", sql, "")
			xml=xml + "</registro>"
			sql.Next
		Wend
		xml=xml + "</registros>"
		Set sql=Nothing
		ServiceVar("psXml") = CStr(xml)
	Else
        Dim viAchou As Integer
		Set sql=NewQuery
		sql.Add("SELECT B.HANDLE, B.BENEFICIARIO, B.NOME, P.DESCRICAO PLANO, M.CARTAONACIONALSAUDE CNS, CART.DATAFINALVALIDADE VALIDADECARTAO, M.SEXO,B.DATACANCELAMENTO")
		sql.Add("FROM SAM_BENEFICIARIO B")
		sql.Add("JOIN SAM_CONTRATO C ON (C.HANDLE=B.CONTRATO)")
		sql.Add("JOIN SAM_PLANO P ON (P.HANDLE=C.PLANO)")
		sql.Add("JOIN SAM_MATRICULA M ON (M.HANDLE=B.MATRICULA)")
		sql.Add("LEFT JOIN SAM_BENEFICIARIO_CARTAOIDENTIF CART ON (CART.BENEFICIARIO=B.HANDLE AND CART.SITUACAO='N')")
		sql.Add("WHERE B.MATRICULAFUNCIONAL=:MATRICULAFUNCIONAL ")
		sql.ParamByName("MATRICULAFUNCIONAL").AsString=psDigitado
		sql.Active=True

       viAchou = -1

		xml="<registros>"
		While Not sql.EOF
		   viAchou = 1

			xml=xml + "<registro>"
			xml=xml + addXMLAtt( "handle", "handle", sql, "")
			xml=xml + addXMLAtt( "nome", "nome", sql, "caption='Nome' width='300'")
			xml=xml + addXMLAtt( "codigo", "beneficiario", sql, "caption='Código' width='120'")
			xml=xml + addXMLAtt( "plano", "plano", sql, "caption='Plano' width='200'")
			xml=xml + addXMLAtt( "cns", "cns", sql, "caption='CNS' width='90'")
			xml=xml + addXMLAtt( "validadeCartao", "validadeCartao", sql, "caption='Validade cartão' width='90'")
			xml=xml + addXMLAtt( "dataCancelamento", "dataCancelamento", sql, "caption='Data de Cancelamento' width='110'")
			xml=xml + addXMLAtt( "sexo", "sexo", sql, "")
			xml=xml + "</registro>"
			sql.Next
		Wend
		xml=xml + "</registros>"
		Set sql=Nothing
		ServiceVar("psXml") = CStr(xml)

		If viAchou < 0 Then
        	Dim especifico As Object
			Set especifico = CreateBennerObject("ESPECIFICO.UESPECIFICO")
			Set sql=NewQuery

			especifico.MPU_BEN_BuscaBeneficiarioCodigoAfinidade( CurrentSystem, _
								     psDigitado, _
								     sql.TQuery)

			xml="<registros>"
			While Not sql.EOF
				xml=xml + "<registro>"
				xml=xml + addXMLAtt( "handle", "handle", sql, "")
				xml=xml + addXMLAtt( "nome", "nome", sql, "caption='Nome' width='300'")
				xml=xml + addXMLAtt( "codigo", "beneficiario", sql, "caption='Código' width='120'")
				xml=xml + addXMLAtt( "plano", "plano", sql, "caption='Plano' width='200'")
				xml=xml + addXMLAtt( "cns", "cns", sql, "caption='CNS' width='90'")
				xml=xml + addXMLAtt( "validadeCartao", "validadeCartao", sql, "caption='Validade cartão' width='90'")
				xml=xml + addXMLAtt( "sexo", "sexo", sql, "")
				xml=xml + "</registro>"
				sql.Next
			Wend
			xml=xml + "</registros>"
			Set sql=Nothing
			ServiceVar("psXml") = CStr(xml)
		End If
	End If

	GoTo fim

	erro:
		ServiceVar("psXml") = CStr(Err.Description)
	fim:
		Set sp=Nothing

End Sub
