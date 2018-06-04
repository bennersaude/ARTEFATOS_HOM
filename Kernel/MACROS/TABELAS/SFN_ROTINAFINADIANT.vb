'HASH: 97D56D19DDD22D558F9A52B58E27ECC9
'#Uses "*bsShowMessage"

'Macro: SFN_ROTINAFINADIANT
'sfn_rotinafinadiant

Option Explicit

Public Sub BOTAOCANCELAR_OnClick()
  Dim SQLRotFin As Object

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição","I")
    Exit Sub
  End If


  If CurrentQuery.FieldByName("SITUACAO").AsString <> "5" Then
    bsShowMessage("A rotina não está processada","I")
    Exit Sub
  End If

  If VisibleMode Then

		Dim vContainer As CSDContainer
		Set vContainer = NewContainer

		vContainer.AddFields("HANDLE:INTEGER;INTERFACE:STRING")

		vContainer.Insert
		vContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		vContainer.Field("INTERFACE").AsString = "SamPagamento.Adiantamento_Cancelar"

		Dim Interface As Object
		Set Interface = CreateBennerObject("BSINTERFACE.Rotinas")
		Interface.Executar(CurrentSystem,vContainer)

		Set Interface = Nothing

  Else

        Dim vsMensagemErro As String
		Dim viRetorno As Long

		Dim qSql As BPesquisa
		Set qSql = NewQuery

		qSql.Add("SELECT CF.COMPETENCIA,                                ")
		qSql.Add("       RF.SEQUENCIA                                   ")
		qSql.Add("  FROM SFN_ROTINAFIN RF                               ")
		qSql.Add("  JOIN SFN_COMPETFIN CF ON (CF.HANDLE = RF.COMPETFIN) ")
		qSql.Add(" WHERE RF.HANDLE = :PHANDLE                           ")
		qSql.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
		qSql.Active  = True


		Dim Obj As Object
	    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
	    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
	                                 "SamPagamento", _
	                                 "Adiantamento_Cancelar", _
	                                 "Faturamento de Adiantamento Competência: " + FormatDateTime2("MM/YYYY",qSql.FieldByName("COMPETENCIA").AsDateTime) & " Seq.: " &  Str(qSql.FieldByName("SEQUENCIA").AsInteger) & " - Cancelamento", _
	                                 CurrentQuery.FieldByName("HANDLE").AsInteger, _
	                                 "SFN_ROTINAFINADIANT", _
	                                 "SITUACAO", _
	                                 "", _
	                                 "", _
	                                 "C", _
	                                 False, _
	                                  vsMensagemErro, _
	                                 Null)


	    If viRetorno = 0 Then
			bsShowMessage("Processo enviado para execução no servidor!", "I")
	 	Else
			bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
	  	End If

	  	Set Obj = Nothing
	  	Set qSql = Nothing

  End If

  WriteAudit("C", HandleOfTable("SFN_ROTINAFINADIANT"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Faturamento de Adiantamentos - Cancelamento")
End Sub

Public Sub BOTAOPROCESSAR_OnClick()

  If CurrentQuery.State <>1 Then
    bsShowMessage("Os parâmetros não podem estar em edição","I")
    Exit Sub
  End If

  Dim qRotinaFin As Object
  Set qRotinaFin = NewQuery

  qRotinaFin.Add("SELECT DATACONTABIL FROM SFN_ROTINAFIN WHERE HANDLE = :PROTINAFIN")
  qRotinaFin.ParamByName("PROTINAFIN").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
  qRotinaFin.Active = True

  Dim vDataContabil As Date
  vDataContabil = qRotinaFin.FieldByName("DATACONTABIL").AsDateTime
  Set qRotinaFin = Nothing

  If CurrentQuery.FieldByName("DATAADIANTAMENTO").AsDateTime < vDataContabil Then
    bsShowMessage("A data de adiantamento não pode ser anterior a data contábil","I")
    Exit Sub
  End If

  If CurrentQuery.FieldByName("DATAADIANTAMENTO").AsDateTime < ServerDate Then
    bsShowMessage("A data de adiantamento não pode ser anterior a hoje","I")
    Exit Sub
  End If

If VisibleMode Then

		Dim vContainer As CSDContainer
		Set vContainer = NewContainer

		vContainer.AddFields("HANDLE:INTEGER;INTERFACE:STRING")
		
		vContainer.Insert
		vContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		vContainer.Field("INTERFACE").AsString = "SamPagamento.Adiantamento_Processar"

		Dim Interface As Object
		Set Interface = CreateBennerObject("BSINTERFACE.Rotinas")
		Interface.Executar(CurrentSystem,vContainer)

		Set Interface = Nothing

  Else

        Dim vsMensagemErro As String
		Dim viRetorno As Long


		Dim qSql As BPesquisa
		Set qSql = NewQuery

		qSql.Add("SELECT CF.COMPETENCIA,                                ")
		qSql.Add("       RF.SEQUENCIA                                   ")
		qSql.Add("  FROM SFN_ROTINAFIN RF                               ")
		qSql.Add("  JOIN SFN_COMPETFIN CF ON (CF.HANDLE = RF.COMPETFIN) ")
		qSql.Add(" WHERE RF.HANDLE = :PHANDLE                           ")
		qSql.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
		qSql.Active  = True



		Dim Obj As Object
	    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
	    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
	                                 "SamPagamento", _
	                                 "Adiantamento_Processar", _
	                                 "Faturamento de Adiantamento Competência: " + FormatDateTime2("MM/YYYY",qSql.FieldByName("COMPETENCIA").AsDateTime) & " Seq.: " &  Str(qSql.FieldByName("SEQUENCIA").AsInteger) & " - Processamento", _
	                                 CurrentQuery.FieldByName("HANDLE").AsInteger, _
	                                 "SFN_ROTINAFINADIANT", _
	                                 "SITUACAO", _
	                                 "", _
	                                 "", _
	                                 "P", _
	                                 False, _
	                                  vsMensagemErro, _
	                                 Null)


	    If viRetorno = 0 Then
			bsShowMessage("Processo enviado para execução no servidor!", "I")
	 	Else
			bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
	  	End If

	  	Set Obj = Nothing
	  	Set qSql = Nothing

  End If

  WriteAudit("P", HandleOfTable("SFN_ROTINAFINADIANT"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Faturamento de Adiantamentos - Processamento")

End Sub

Public Sub PEGFINAL_OnPopup(ShowPopup As Boolean)

  Dim datapag As Date
  Dim STRX As String

  datapag = CurrentQuery.FieldByName("DATAADIANTAMENTO").AsDateTime

  STRX = " AND SAM_PEG.DATAADIANTAMENTO = "+ SQLDate(datapag)

  If WebMode Then
	  PEGFINAL.WebLocalWhere = "SAM_PEG.PEG >= " + _
	                        "(Select PEG FROM SAM_PEG WHERE SAM_PEG.HANDLE = " + _
	                        Str(CurrentQuery.FieldByName("PEGINICIAL").AsInteger) + STRX + ")"
  Else
	  PEGFINAL.LocalWhere = "SAM_PEG.PEG >= " + _
	                        "(Select PEG FROM SAM_PEG WHERE SAM_PEG.HANDLE = " + _
	                        Str(CurrentQuery.FieldByName("PEGINICIAL").AsInteger) + STRX + ")"

  End If

End Sub

Public Sub PEGINICIAL_OnPopup(ShowPopup As Boolean)

  Dim datapag As Date
  Dim STRX As String
  datapag = CurrentQuery.FieldByName("DATAADIANTAMENTO").AsDateTime
  STRX = "SAM_PEG.DATAADIANTAMENTO = "+ SQLDate(datapag)

  If WebMode Then
	  PEGINICIAL.WebLocalWhere = STRX
  Else
      PEGINICIAL.LocalWhere = STRX
  End If

End Sub

Public Sub TABLE_AfterScroll()

	BOTAOPROCESSAR.Enabled = CurrentQuery.FieldByName("SITUACAO").AsString = "1"
	BOTAOCANCELAR.Enabled = CurrentQuery.FieldByName("SITUACAO").AsString = "5"

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  If CurrentQuery.FieldByName("SITUACAO").Value <> "1" Then
    bsShowMessage("Alteração não permitida. A rotina não está Aberta.","E")
    CanContinue = False
  End If

End Sub


Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "PROCESSAR"
			BOTAOPROCESSAR_OnClick
		Case "CANCELAR"
			BOTAOCANCELAR_OnClick
	End Select
End Sub
