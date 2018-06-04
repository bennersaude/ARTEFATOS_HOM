'HASH: 3D891CFD9FB54934733870120B0DE041
Option Explicit

Public Sub Main()
	Dim qSelecionaRegistros	 As Object
	Dim qReservaRegistro	 As Object
	Dim qLiberaRegistro	     As Object
	Dim qLiberaRegistroErro  As Object
	Dim qPreChecaRegistro	 As Object
	Dim qLiberaControle		 As Object
	Dim vIProcessando 		 As Long
	Dim vDLLDireciona		 As Object
	Set qSelecionaRegistros = NewQuery
	Set qReservaRegistro = NewQuery
	Set qLiberaRegistro = NewQuery
	Set qLiberaRegistroErro = NewQuery
	Set qPreChecaRegistro = NewQuery
	Set qLiberaControle = NewQuery

	Dim qGuia               As BPesquisa
	Dim qSituacaoRecurso    As BPesquisa
	Dim qDeletePEG          As BPesquisa
	Dim qUpdMensagemTiss    As BPesquisa

	Set qGuia            = NewQuery
    Set qSituacaoRecurso = NewQuery
	Set qDeletePEG       = NewQuery
	Set qUpdMensagemTiss = NewQuery

	vIProcessando = NewHandle("TMP_AUX1")


	' reserva todos os registros para que outro agendamento não consiga e não possa selecionar o mesmo registro,
	' pois o CONTROLE vIProcessando é único portando cada execução do agendamento será uma "Rotina a parte"
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
	qReservaRegistro.Add("	 AND HANDLE IN (SELECT MENSAGEMTISS                                                                   ")
  	qReservaRegistro.Add("		              FROM TIS_RECURSOGLOSA)        												      ")


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


	' query para somente certificar que o registro atual está e será processado somente por esta rotina
	qPreChecaRegistro.Clear
	qPreChecaRegistro.Add("SELECT CONTROLE                                      ")
	qPreChecaRegistro.Add("  FROM SAM_PRESTADOR_MENSAGEMTISS                    ")
	qPreChecaRegistro.Add(" WHERE HANDLE = :HANDLE                              ")

	' query liberar o registro que acabou de ser identificado
	qLiberaRegistro.Clear
	qLiberaRegistro.Add("UPDATE SAM_PRESTADOR_MENSAGEMTISS                      ")
	qLiberaRegistro.Add("   SET SITUACAO = :SITUACAOPROCESSADO                  ")

	If (InStr(UCase(SQLServer), "MSSQL") > 0) Then
		qLiberaRegistro.Add("  FROM SAM_PRESTADOR_MENSAGEMTISS  READPAST          ")
	End If

	qLiberaRegistro.Add(" WHERE CONTROLE = :CONTROLE                            ")
	qLiberaRegistro.Add("   AND HANDLE   = :HANDLE                              ")

	' query para setar registro com erro

	qLiberaRegistroErro.Clear
	qLiberaRegistroErro.Add("UPDATE SAM_PRESTADOR_MENSAGEMTISS                  ")
	qLiberaRegistroErro.Add("   SET SITUACAO = :SITUACAOERRO,                   ")
	qLiberaRegistroErro.Add("       CONTROLE = NULL,                            ")
	qLiberaRegistroErro.Add("       (OCORRENCIAS = :OCORRENCIAS)                ")

	If (InStr(UCase(SQLServer), "MSSQL") > 0) Then
		qLiberaRegistroErro.Add("  FROM SAM_PRESTADOR_MENSAGEMTISS  READPAST      ")
	End If

	qLiberaRegistroErro.Add(" WHERE HANDLE   = :HANDLE                          ")
	qLiberaRegistroErro.Add("   AND CONTROLE = :CONTROLE                        ")


	Dim vDLLImportar As Object
	Set vDLLImportar = CreateBennerObject("Benner.Saude.WSTiss.PreVersionador.PreVersionador")

	Dim vRetornoImportacao As String

	qLiberaControle.Clear
	qLiberaControle.Add("UPDATE SAM_PRESTADOR_MENSAGEMTISS              ")
	qLiberaControle.Add("   SET CONTROLE = NULL,                        ")
	qLiberaControle.Add("       SITUACAO = 'E',                         ")
	qLiberaControle.Add("       OCORRENCIAS = (:OCORRENCIAS)            ")

	If (InStr(UCase(SQLServer), "MSSQL") > 0) Then
		qLiberaControle.Add("  FROM SAM_PRESTADOR_MENSAGEMTISS  READPAST   ")
	End If

	qLiberaControle.Add(" WHERE HANDLE = :HANDLE                         ")
	qLiberaControle.Add("   AND CONTROLE = :CONTROLE                     ")

    qSituacaoRecurso.Active = False
    qSituacaoRecurso.Clear
    qSituacaoRecurso.Add("SELECT SITUACAO                      ")
    qSituacaoRecurso.Add("  FROM TIS_RECURSOGLOSA              ")
    qSituacaoRecurso.Add(" WHERE MENSAGEMTISS = :HMENSAGEMTISS ")

    qGuia.Active = False
    qGuia.Clear
	qGuia.Add("Select COUNT(HANDLE) QTDE ")
    qGuia.Add("  FROM SAM_GUIA G         ")
	qGuia.Add(" WHERE PEG = :HPEG        ")

    qDeletePEG.Active = False
	qDeletePEG.Clear
	qDeletePEG.Add("DELETE FROM SAM_PEG WHERE HANDLE = :HPEG")

	qUpdMensagemTiss.Active = False
	qUpdMensagemTiss.Clear
	qUpdMensagemTiss.Add("UPDATE SAM_PRESTADOR_MENSAGEMTISS ")
	qUpdMensagemTiss.Add("   SET HANDLEPEG = NULL,          ")
	qUpdMensagemTiss.Add("       NUMEROPROTOCOLO = NULL,    ")
	qUpdMensagemTiss.Add("       QTDGUIASIMPORTADAS = NULL, ")
	qUpdMensagemTiss.Add("       TIPOGUIASIMPORTADAS = NULL,")
	qUpdMensagemTiss.Add("       VALORGUIASIMPORTADAS = NULL")
	qUpdMensagemTiss.Add(" WHERE HANDLE    = :HMENSAGEM     ")



	While Not qSelecionaRegistros.EOF

		On Error GoTo erro

'		qPreChecaRegistro.Active = False
'		qPreChecaRegistro.ParamByName("HANDLE").AsInteger = qSelecionaRegistros.FieldByName("HANDLE").AsInteger
'		qPreChecaRegistro.Active = True
'
'		If qPreChecaRegistro.FieldByName("CONTROLE").AsInteger = 0 Then

			'StartTransaction

			vRetornoImportacao = ""


			qPreChecaRegistro.Active = False
			qPreChecaRegistro.ParamByName("HANDLE").AsInteger = qSelecionaRegistros.FieldByName("HANDLE").AsInteger
			qPreChecaRegistro.Active = True

			If vIProcessando = qPreChecaRegistro.FieldByName("CONTROLE").AsInteger Then ' Se entrar aqui é pq o registro foi reservado por esta execução do agendamento

				If Not vDLLImportar.Exec(CurrentSystem, qSelecionaRegistros.FieldByName("HANDLE").AsInteger, vRetornoImportacao) Then

					qLiberaControle.Active = False
					qLiberaControle.ParamByName("HANDLE").AsInteger = qSelecionaRegistros.FieldByName("HANDLE").AsInteger
					qLiberaControle.ParamByName("CONTROLE").AsInteger = vIProcessando
					qLiberaControle.ParamByName("OCORRENCIAS").AsString = vRetornoImportacao
					qLiberaControle.ExecSQL
				Else
				    qSituacaoRecurso.Active = False
				    qSituacaoRecurso.ParamByName("HMENSAGEMTISS").AsInteger = qSelecionaRegistros.FieldByName("HANDLE").AsInteger
				    qSituacaoRecurso.Active = True

				    If qSituacaoRecurso.FieldByName("SITUACAO").AsString = "4" Then
                      qGuia.Active = False
	  				  qGuia.ParamByName("HPEG").AsInteger = qSelecionaRegistros.FieldByName("HANDLEPEG").AsInteger
					  qGuia.Active = True

					  If qGuia.FieldByName("QTDE").AsInteger = 0 Then
					    qDeletePEG.Active = False
					    qDeletePEG.ParamByName("HPEG").AsInteger = qSelecionaRegistros.FieldByName("HANDLEPEG").AsInteger
					    qDeletePEG.ExecSQL

					    qUpdMensagemTiss.Active = False
					    qUpdMensagemTiss.ParamByName("HMENSAGEM").AsInteger = qSelecionaRegistros.FieldByName("HANDLE").AsInteger
					    qUpdMensagemTiss.ExecSQL
					  End If
				    End If

					' Colocando o registro como processado
					qLiberaRegistro.Active = False
					qLiberaRegistro.ParamByName("CONTROLE").AsInteger = vIProcessando
					qLiberaRegistro.ParamByName("HANDLE").AsInteger = qSelecionaRegistros.FieldByName("HANDLE").AsInteger
					qLiberaRegistro.ParamByName("SITUACAOPROCESSADO").AsString = "P"
					qLiberaRegistro.ExecSQL

				End If

			End If

			'Commit

'		End If

		GoTo ProximoRegistro

		erro : ' caso ocorra erro, o registro atual voltará como liberado ou seja situacao = 'A' e poderá ser pego em outro agendamento e vai pro próximo registro (mensagemtiss)

			qLiberaRegistroErro.Active = False
			qLiberaRegistroErro.ParamByName("HANDLE").AsInteger = qSelecionaRegistros.FieldByName("HANDLE").AsInteger
			qLiberaRegistroErro.ParamByName("CONTROLE").AsInteger = vIProcessando
			qLiberaRegistroErro.ParamByName("SITUACAOERRO").AsString = "E"
			qLiberaRegistroErro.ParamByName("OCORRENCIAS").AsString = Err.Description
			qLiberaRegistroErro.ExecSQL

			'Commit

		ProximoRegistro :
			qSelecionaRegistros.Next
	Wend

	Set vDLLImportar        = Nothing
	Set qSelecionaRegistros = Nothing
	Set qReservaRegistro    = Nothing
	Set qLiberaRegistro     = Nothing
	Set qLiberaRegistroErro = Nothing
	Set qPreChecaRegistro   = Nothing
	Set qGuia               = Nothing
    Set qSituacaoRecurso    = Nothing
	Set qDeletePEG          = Nothing
	Set qUpdMensagemTiss    = Nothing
End Sub
