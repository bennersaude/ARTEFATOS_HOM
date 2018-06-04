'HASH: C1185826CA0EFFC6848BF3E18FB64B6B

'#Uses "*addXMLAtt"

Sub Main

	Dim psDigitado As String
	Dim handle As Long

	psDigitado = CStr( ServiceVar("psDigitado") )

	Dim sp As BStoredProc
	Set sp=NewStoredProc

	On Error GoTo erro


	sp.Name="BSTISS_VALIDARPRESTADOR"
	sp.AddParam("P_DIGITADO",ptInput, ftString, 100)
	sp.AddParam("P_ORIGEM",ptInput, ftString, 1)
	sp.AddParam("P_HANDLE",ptOutput, ftInteger)
	sp.ParamByName("P_DIGITADO").AsString = psDigitado
	sp.ParamByName("P_ORIGEM").AsString= "0" 'origem zero procura por cpf, cnpj, etc..
	sp.ExecProc
	handle = sp.ParamByName("P_HANDLE").AsInteger

	If handle > 0 Then
		Dim sql As BPesquisa
		Set sql=NewQuery
		sql.Add("SELECT P.HANDLE, P.FISICAJURIDICA, P.PRESTADOR, P.NOME, CONSELHO.SIGLA CONSELHOSIGLA, P.INSCRICAOCR, UF2.SIGLA UFCONSELHO, CBOS.ESTRUTURA CBOS")
		sql.Add("FROM SAM_PRESTADOR P")
		sql.Add("LEFT JOIN SAM_CONSELHO CONSELHO ON (CONSELHO.HANDLE=P.CONSELHOREGIONAL)")
		sql.Add("LEFT JOIN ESTADOS UF2 ON (UF2.HANDLE=P.UFCR)")
		sql.Add("LEFT JOIN SAM_CBO CBOS ON (CBOS.HANDLE=P.CBO)")
		sql.Add("WHERE P.HANDLE=:HANDLE ")
		sql.ParamByName("HANDLE").AsInteger=handle
		sql.Active=True

		Dim xml As String
		xml="<registros>"
		While Not sql.EOF
			xml=xml + "<registro>"
			xml=xml + addXMLAtt( "handle", "handle", sql, "")
			xml=xml + addXMLAtt( "nome", "nome", sql, "caption='Nome' width='300'")
			xml=xml + addXMLAtt( "codigo", "prestador", sql, "caption='Código' width='120'")
			xml=xml + addXMLAtt( "fisicaJuridica", "fisicaJuridica", sql, "")
			xml=xml + addXMLAtt( "conselhosigla", "conselhoSigla", sql, " ")
			xml=xml + addXMLAtt( "inscricaocr", "inscricaocr", sql, " ")
			xml=xml + addXMLAtt( "ufconselho", "ufConselho", sql, " ")
			xml=xml + addXMLAtt( "cbos", "cbos", sql, "caption='' ")
			xml=xml + "</registro>"
			sql.Next
			If (piHandle>0) Then
			  Exit While
			End If

		Wend
		xml=xml + "</registros>"
		Set sql=Nothing
		ServiceVar("psResult") = CStr(xml)
	End If

	GoTo fim

	erro:

		ServiceVar("psResult") = CStr(Err.Description)

	fim:
		Set sp=Nothing

End Sub
