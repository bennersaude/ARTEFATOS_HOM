'HASH: E7C9AE98C74346E500CB2C9E454913CA

Public Sub Main
	Dim pNomeTabela As String 'Parametro de entrada
	Dim pCaminhoArquivo As String 'Parametro de entrada
	Dim pNomeArquivo As String 'Parametro de entrada
	Dim pCampoFK As String 'Parametro de entrada

	Dim pHandleMensagemTISS As Long ' Retorno para c#

	pNomeTabela       = ServiceVar("pNomeTabela")                ' "SAM_PRESTADOR_VALIDAMSGTISS"
	pCaminhoArquivo   = ServiceVar("pCaminho")                   ' "C:\"
	pCampoNomeArquivo = ServiceVar("pCampoNomeArquivo")          ' "ARQUIVOVALIDADO"
	pNomeArquivo      = ServiceVar("pNomeArquivo")               ' "2_d1ac378d6e93796aeb39ff4307daed5d.xml"
	pCampoNomeFK      = ServiceVar("pCampoNomeFK")               ' "PRESTADOR"
	pValorCampoFK     = ServiceVar("pValorCampoFK")              ' 555

	Dim qLocalizaRecebedor As Object
	Set qLocalizaRecebedor = NewQuery

	Dim qLocalizaArquivo As Object
	Set qLocalizaArquivo = NewQuery
	qLocalizaArquivo.Clear
	qLocalizaArquivo.Add("SELECT HANDLE                      ")
	qLocalizaArquivo.Add("  FROM " + pNomeTabela )
	qLocalizaArquivo.Add(" WHERE " + pCampoNomeArquivo + " = :ARQUIVO  ")
	qLocalizaArquivo.Add("   AND " + pCampoNomeFK      + " = :CAMPOFK")
	qLocalizaArquivo.ParamByName("CAMPOFK").AsInteger = pValorcampoFK
	qLocalizaArquivo.ParamByName("ARQUIVO").AsString = pNomeArquivo
	qLocalizaArquivo.Active = True
	If qLocalizaArquivo.FieldByName("HANDLE").AsInteger > 0 Then
		pHandleMensagemTISS = qLocalizaArquivo.FieldByName("HANDLE").AsInteger
	Else
		Dim qInsereMensagemTISS As Object
		Set qInsereMensagemTISS = NewQuery

		pHandleMensagemTISS = NewHandle(pNomeTabela)
		qInsereMensagemTISS.Clear
		qInsereMensagemTISS.Add("INSERT INTO " + pNomeTabela + " (HANDLE , " + pCampoNomeFK + ") ")
		qInsereMensagemTISS.Add("                         VALUES (:HANDLE, :pCampoNomeFK) ")
		qInsereMensagemTISS.ParamByName("HANDLE").AsInteger = pHandleMensagemTISS
		qInsereMensagemTISS.ParamByName("pCampoNomeFK").AsInteger = pValorcampoFK
		qInsereMensagemTISS.ExecSQL
		SetFieldDocument(pNomeTabela, pCampoNomeArquivo, pHandleMensagemTISS, pCaminhoArquivo + pNomeArquivo, True)
		Set qInsereMensagemTISS = Nothing
	End If
	Set qLocalizaArquivo = Nothing
	Set qLocalizaRecebedor = Nothing

	ServiceResult = CLng(pHandleMensagemTISS)


End Sub
