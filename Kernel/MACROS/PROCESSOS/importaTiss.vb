'HASH: 338052E2B49A916403AA81330284D0A5
Public Sub Main()
	Dim qSelecionaRegistros	 As Object
	Dim qReservaRegistro	 As Object
	Dim qLiberaRegistro	     As Object
	Dim qLiberaRegistroErro  As Object
	Dim qPreChecaRegistro	 As Object
	Dim qPreencheSituacaoTiss As Object
	Dim qSituacaoTiss 		As Object
	Dim vIProcessando 		 As Long
	Dim vDLLDireciona		 As Object
	Set qSelecionaRegistros = NewQuery
	Set qReservaRegistro = NewQuery
	Set qLiberaRegistro = NewQuery
	Set qLiberaRegistroErro = NewQuery
	Set qPreChecaRegistro = NewQuery
	Set qSituacaoTiss = NewQuery
	Set qPreencheSituacaoTiss = NewQuery

	vIProcessando = NewHandle("TMP_AUX1")


	' reserva todos os registros para que outro agendamento nÃ£o consiga e nÃ£o possa selecionar o mesmo registro,
	' pois o CONTROLE vIProcessando Ã© Ãºnico portando cada execuÃ§Ã£o do agendamento serÃ¡ uma "Rotina a parte"
	qReservaRegistro.Clear
	qReservaRegistro.Add("UPDATE SAM_PRESTADOR_MENSAGEMTISS                     ")
	qReservaRegistro.Add("   SET CONTROLE = :CONTROLE,                          ")
	qReservaRegistro.Add("       SITUACAO = :SITUACAO                           ")

	If (InStr(UCase(SQLServer), "MSSQL")) Then
		qReservaRegistro.Add("  FROM SAM_PRESTADOR_MENSAGEMTISS  READPAST           ")
	End If

	qReservaRegistro.Add(" WHERE CONTROLE IS NULL AND SITUACAO = :SITUACAOABERTA AND HANDLEPEG > 0 AND ARQUIVOAGENDADO IS NOT NULL")
	qReservaRegistro.Add("	 AND HANDLE NOT IN (SELECT MENSAGEMTISS                                                               ")
  	qReservaRegistro.Add("		                  FROM TIS_IMPORTACAOXMLLOTE_ARQ ARQ                                              ")
  	qReservaRegistro.Add("		                  JOIN TIS_IMPORTACAOXMLLOTE     LOT ON (LOT.HANDLE = ARQ.IMPORTACAOXMLLOTE))     ")
	qReservaRegistro.Add("	 AND HANDLE NOT IN (SELECT MENSAGEMTISS                                                               ")
  	qReservaRegistro.Add("		                  FROM TIS_RECURSOGLOSA)        												  ")


	If SessionVar("PRIORIDADE") <> "" Then
		qReservaRegistro.Add("  AND PRIORIDADE = :PRIORIDADE                 ")
		qReservaRegistro.ParamByName("PRIORIDADE").AsInteger = CLng(SessionVar("PRIORIDADE"))
	End If

	qReservaRegistro.Active = False
	qReservaRegistro.ParamByName("CONTROLE").AsInteger = vIProcessando
	qReservaRegistro.ParamByName("SITUACAO").AsString = "S"
	qReservaRegistro.ParamByName("SITUACAOABERTA").AsString = "A"
	qReservaRegistro.ExecSQL


	qSelecionaRegistros.Clear
	qSelecionaRegistros.Add("SELECT *                                           ")
	qSelecionaRegistros.Add("  FROM SAM_PRESTADOR_MENSAGEMTISS                  ")
		qSelecionaRegistros.Add(" WHERE SITUACAO = :SITUACAO AND CONTROLE = :CONTROLE ")
	If SessionVar("PRIORIDADE") <> "" Then
		qSelecionaRegistros.Add("  AND PRIORIDADE = :PRIORIDADE                 ")
		qSelecionaRegistros.ParamByName("PRIORIDADE").AsInteger = CLng(SessionVar("PRIORIDADE"))
	End If
	qSelecionaRegistros.Add(" ORDER BY HANDLE ASC                               ")
	qSelecionaRegistros.ParamByName("CONTROLE").AsInteger = vIProcessando
	qSelecionaRegistros.ParamByName("SITUACAO").AsString = "S"
	qSelecionaRegistros.Active = True


	' query para somente certificar que o registro atual estÃ¡ e serÃ¡ processado somente por esta rotina
	qPreChecaRegistro.Clear
	qPreChecaRegistro.Add("SELECT CONTROLE                                      ")
	qPreChecaRegistro.Add("  FROM SAM_PRESTADOR_MENSAGEMTISS                    ")
	qPreChecaRegistro.Add(" WHERE HANDLE = :HANDLE                              ")

	' query liberar o registro que acabou de ser identificado
	qLiberaRegistro.Clear
	qLiberaRegistro.Add("UPDATE SAM_PRESTADOR_MENSAGEMTISS                      ")
	qLiberaRegistro.Add("   SET SITUACAO = :SITUACAOPROCESSADO                  ")

	If (InStr(UCase(SQLServer), "MSSQL") > 0) Then
		qLiberaRegistro.Add("  FROM SAM_PRESTADOR_MENSAGEMTISS  READPAST        ")
	End If

	qLiberaRegistro.Add(" WHERE CONTROLE = :CONTROLE                            ")
	qLiberaRegistro.Add("   AND HANDLE   = :HANDLE                              ")

	' query para setar registro com erro

	qLiberaRegistroErro.Clear
	qLiberaRegistroErro.Add(" UPDATE SAM_PRESTADOR_MENSAGEMTISS   ")
	qLiberaRegistroErro.Add("    SET SITUACAO = :SITUACAOERRO,    ")
	qLiberaRegistroErro.Add("        CONTROLE = NULL,             ")
	qLiberaRegistroErro.Add("        OCORRENCIAS = :OCORRENCIAS   ")

	If (InStr(UCase(SQLServer), "MSSQL") > 0) Then
		qLiberaRegistroErro.Add(" FROM SAM_PRESTADOR_MENSAGEMTISS  READPAST ")
	End If

	qLiberaRegistroErro.Add(" WHERE CONTROLE = :CONTROLE          ")
	qLiberaRegistroErro.Add("   AND HANDLE   = :HANDLE            ")

	Dim vDLLImportar As Object
	Set vDLLImportar = CreateBennerObject("Benner.Saude.WSTiss.PreVersionador.PreVersionador")

	Dim vRetornoImportacao As String


	While Not qSelecionaRegistros.EOF

		On Error GoTo erro

			vRetornoImportacao = ""

			qPreChecaRegistro.Active = False
			qPreChecaRegistro.ParamByName("HANDLE").AsInteger = qSelecionaRegistros.FieldByName("HANDLE").AsInteger
			qPreChecaRegistro.Active = True

			If vIProcessando = qPreChecaRegistro.FieldByName("CONTROLE").AsInteger Then ' Se entrar aqui Ã© pq o registro foi reservado por esta execuÃ§Ã£o do agendamento
				If Not vDLLImportar.Exec(CurrentSystem, qSelecionaRegistros.FieldByName("HANDLE").AsInteger, vRetornoImportacao) Then
					qLiberaRegistroErro.Active = False
					qLiberaRegistroErro.ParamByName("CONTROLE").AsInteger = vIProcessando
					qLiberaRegistroErro.ParamByName("HANDLE").AsInteger = qSelecionaRegistros.FieldByName("HANDLE").AsInteger
					qLiberaRegistroErro.ParamByName("SITUACAOERRO").AsString = "E"
					qLiberaRegistroErro.ParamByName("OCORRENCIAS").AsString = vRetornoImportacao
					qLiberaRegistroErro.ExecSQL

				Else
					' Colocando o registro como processado
					qLiberaRegistro.Active = False
					qLiberaRegistro.ParamByName("CONTROLE").AsInteger = vIProcessando
					qLiberaRegistro.ParamByName("HANDLE").AsInteger = qSelecionaRegistros.FieldByName("HANDLE").AsInteger
					qLiberaRegistro.ParamByName("SITUACAOPROCESSADO").AsString = "P"
					qLiberaRegistro.ExecSQL
				End If
			End If

		GoTo ProximoRegistro

		erro : ' caso ocorra erro, o registro atual voltarÃ¡ como liberado ou seja situacao = 'A' e poderÃ¡ ser pego em outro agendamento e vai pro prÃ³ximo registro (mensagemtiss)

			qLiberaRegistroErro.Active = False
			qLiberaRegistroErro.ParamByName("CONTROLE").AsInteger = vIProcessando
			qLiberaRegistroErro.ParamByName("HANDLE").AsInteger = qSelecionaRegistros.FieldByName("HANDLE").AsInteger
			qLiberaRegistroErro.ParamByName("SITUACAOERRO").AsString = "E"
			qLiberaRegistroErro.ParamByName("OCORRENCIAS").AsString = Err.Description
			qLiberaRegistroErro.ExecSQL
		ProximoRegistro :
			qSelecionaRegistros.Next
	Wend

	qSituacaoTiss.Clear
	qSituacaoTiss.Add("SELECT TV.VERSAO FROM SAM_PRESTADOR_MENSAGEMTISS MT")
	qSituacaoTiss.Add("  JOIN SAM_PEG    P  ON MT.HANDLEPEG = P.HANDLE")
	qSituacaoTiss.Add("  JOIN TIS_VERSAO TV ON TV.HANDLE = P.VERSAOTISS")
	qSituacaoTiss.Add(" WHERE MT.HANDLE = :pHANDLE")
	qSituacaoTiss.ParamByName("pHANDLE").AsInteger = qSelecionaRegistros.FieldByName("HANDLE").AsInteger
	qSituacaoTiss.Active = True

	If (Not qSituacaoTiss.EOF) Then
		qPreencheSituacaoTiss.Clear
		qPreencheSituacaoTiss.Add("UPDATE SAM_PRESTADOR_MENSAGEMTISS")
		qPreencheSituacaoTiss.Add("   SET VERSAOTISS =:VERSAOTISS")
		qPreencheSituacaoTiss.Add(" WHERE HANDLE = :pHANDLEV")
		qPreencheSituacaoTiss.ParamByName("pHANDLEV").AsInteger = qSelecionaRegistros.FieldByName("HANDLE").AsInteger
		qPreencheSituacaoTiss.ParamByName("VERSAOTISS").AsString = qSituacaoTiss.FieldByName("VERSAO").AsString
		qPreencheSituacaoTiss.ExecSQL
	End If

	Set vDLLImportar = Nothing
	Set qSelecionaRegistros = Nothing
	Set qReservaRegistro = Nothing
	Set qLiberaRegistro = Nothing
	Set qLiberaRegistroErro = Nothing
	Set qPreChecaRegistro = Nothing

	Set qSituacaoTiss = Nothing
	Set qPreencheSituacaoTiss = Nothing

End Sub
