'HASH: 2354C75F8F26C5074509A1E731508DC5
'Macro: SFN_ROTINAFINFAT
'A funcao NodeInternalCode é utilizada para determinar se a carga correspondente é da Tarefas de Modelo,
'sendo, mostra o Tab - Modelo para agendamento, não sendo, mostra o Tab - Rotina
'Alteração: 26/12/2005
'      SMS: 52120 - Marcelo Barbosa
'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELARPARCELAMENTO_OnClick()
  If CurrentQuery.State <> 1 Then
		bsShowMessage("Os parâmetros não podem estar em edição", "I")
		Exit Sub
  End If

  Dim Obj As Object

  If VisibleMode Then
    Set Obj = CreateBennerObject("BSINTERFACE0016.RotinaFaturamentoBeneficiarios")
    Obj.CancelarParcelamento(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "T", 0)
  Else
    Dim SQL As Object
    Set SQL = NewQuery

    SQL.Clear
    SQL.Add("SELECT SFAT.DESCRICAO DESCRICAOTIPOFATURAMENTO,")
    SQL.Add("       CFIN.COMPETENCIA,")
    SQL.Add("       RFIN.SEQUENCIA")
    SQL.Add("FROM SFN_ROTINAFIN       RFIN")
    SQL.Add("JOIN SFN_COMPETFIN       CFIN ON RFIN.COMPETFIN       = CFIN.HANDLE")
    SQL.Add("JOIN SIS_TIPOFATURAMENTO SFAT ON CFIN.TIPOFATURAMENTO = SFAT.HANDLE")
    SQL.Add("WHERE RFIN.HANDLE = :HROTINAFIN")
    SQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
    SQL.Active = True

    Dim vsMensagemErro As String
    Dim viRetorno As Long
    Dim vcContainer As CSDContainer
    Set vcContainer = NewContainer

    vcContainer.AddFields("HANDLE:INTEGER;OPCAOCANCELAMENTO:STRING;HOPCAOCANCELAMENTO:INTEGER")
    vcContainer.Insert
    vcContainer.Field("HANDLE").AsInteger             = CurrentQuery.FieldByName("HANDLE").AsInteger
    vcContainer.Field("OPCAOCANCELAMENTO").AsString   = "T"
    vcContainer.Field("HOPCAOCANCELAMENTO").AsInteger = 0

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBen018", _
                                     "RotinaParcelamentoBeneficiarios_Cancelar", _
                                     "Rotina de Parcelamento de Beneficiários (Cancelar) -" + _
                                       " Faturamento: " + SQL.FieldByName("DESCRICAOTIPOFATURAMENTO").AsString + _
                                       " Competência: " + Str(Format(SQL.FieldByName("COMPETENCIA").AsDateTime, "mm/yyyy")) + _
                                       " Sequência: "   + SQL.FieldByName("SEQUENCIA").AsString, _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_ROTINAFINFAT", _
                                     "SITUACAOPARCELAMENTO", _
                                     "", _
                                     "", _
                                     "C", _
                                     False, _
                                     vsMensagemErro, _
                                     vcContainer)

    Set SQL = Nothing
    Set vcContainer = Nothing

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If
  End If

  Set Obj = Nothing
  If VisibleMode Then
    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If
End Sub

Public Sub BOTAOPARCELAR_OnClick()
  Dim Obj As Object
  Dim vCompetencia As Date

  If CurrentQuery.State <> 1 Then
		bsShowMessage("Os parâmetros não podem estar em edição", "I")
		Exit Sub
  End If

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT CFIN.COMPETENCIA")
  SQL.Add("FROM SFN_COMPETFIN CFIN, SFN_ROTINAFIN RFIN")
  SQL.Add("WHERE RFIN.HANDLE = :HROTINAFIN")
  SQL.Add("  AND CFIN.HANDLE = RFIN.COMPETFIN")

  SQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
  SQL.Active = True

  vCompetencia = SQL.FieldByName("COMPETENCIA").AsDateTime

  SQL.Clear

  SQL.Add("SELECT DATAFATURAMENTOINICIAL, DATAFATURAMENTOFINAL")
  SQL.Add("FROM SFN_PARAMETROSFIN")

  SQL.Active = True

  If Not(SQL.FieldByName("DATAFATURAMENTOINICIAL").IsNull) And _
		 (vCompetencia < SQL.FieldByName("DATAFATURAMENTOINICIAL").AsDateTime) Then
		bsShowMessage("A competência de faturamento é inferior ao período permitido", "I")
		Set SQL = Nothing
		Exit Sub
  End If

  If Not(SQL.FieldByName("DATAFATURAMENTOFINAL").IsNull) And _
		 (vCompetencia > SQL.FieldByName("DATAFATURAMENTOFINAL").AsDateTime) Then
		bsShowMessage("A competência de faturamento é superior ao período permitido", "I")
		Set SQL = Nothing
		Exit Sub
  End If

  If VisibleMode Then
    Set Obj = CreateBennerObject("BSINTERFACE0016.RotinaFaturamentoBeneficiarios")
    Obj.ProcessarParcelamento(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Else
    SQL.Clear
    SQL.Add("SELECT SFAT.DESCRICAO DESCRICAOTIPOFATURAMENTO,")
    SQL.Add("       CFIN.COMPETENCIA,")
    SQL.Add("       RFIN.SEQUENCIA")
    SQL.Add("FROM SFN_ROTINAFIN       RFIN")
    SQL.Add("JOIN SFN_COMPETFIN       CFIN ON RFIN.COMPETFIN       = CFIN.HANDLE")
    SQL.Add("JOIN SIS_TIPOFATURAMENTO SFAT ON CFIN.TIPOFATURAMENTO = SFAT.HANDLE")
    SQL.Add("WHERE RFIN.HANDLE = :HROTINAFIN")
    SQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
    SQL.Active = True

    Dim vsMensagemErro As String
    Dim viRetorno As Long
    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBen018", _
                                     "RotinaParcelamentoBeneficiarios_Processar", _
                                     "Rotina de Parcelamento de Beneficiários (Processar) -" + _
                                       " Faturamento: " + SQL.FieldByName("DESCRICAOTIPOFATURAMENTO").AsString + _
                                       " Competência: " + Str(Format(SQL.FieldByName("COMPETENCIA").AsDateTime, "mm/yyyy")) + _
                                       " Sequência: "   + SQL.FieldByName("SEQUENCIA").AsString, _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SFN_ROTINAFINFAT", _
                                     "SITUACAOPARCELAMENTO", _
                                     "SITUACAOFATURAMENTO", _
                                     "Faturamento não foi processado.", _
                                     "P", _
                                     False, _
                                     vsMensagemErro, _
                                     Null)
    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If
  End If

  Set Obj = Nothing
  Set SQL = Nothing

  If VisibleMode Then
    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim Obj As Object
  Dim vCompetencia As Date

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição", "I")
	Exit Sub
  End If

  Dim SQL As Object
  Set SQL = NewQuery


  'Se não for Apropriação de faturamento antecipado
  If (CurrentQuery.FieldByName("TABTIPOPROCESSO").AsInteger <> 5) Then
    SQL.Add("SELECT CFIN.COMPETENCIA")
    SQL.Add("FROM SFN_COMPETFIN CFIN, SFN_ROTINAFIN RFIN")
    SQL.Add("WHERE RFIN.HANDLE = :HROTINAFIN")
    SQL.Add("  AND CFIN.HANDLE = RFIN.COMPETFIN")

    SQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
    SQL.Active = True

    vCompetencia = SQL.FieldByName("COMPETENCIA").AsDateTime

    SQL.Clear
    SQL.Add("SELECT DATAFATURAMENTOINICIAL, DATAFATURAMENTOFINAL")
    SQL.Add("FROM SFN_PARAMETROSFIN")
    SQL.Active = True

    If Not(SQL.FieldByName("DATAFATURAMENTOINICIAL").IsNull) And _
  		(vCompetencia < SQL.FieldByName("DATAFATURAMENTOINICIAL").AsDateTime) Then
  		bsShowMessage("A competência de faturamento é inferior ao período permitido", "I")
		Set SQL = Nothing
		Exit Sub
    End If

    If Not(SQL.FieldByName("DATAFATURAMENTOFINAL").IsNull) And _
		 (vCompetencia > SQL.FieldByName("DATAFATURAMENTOFINAL").AsDateTime) Then
		bsShowMessage("A competência de faturamento é superior ao período permitido", "I")
		Set SQL = Nothing
		Exit Sub
    End If
  End If

  SQL.Clear
  SQL.Add("SELECT HANDLE")
  SQL.Add("FROM SFN_ROTINAFINFAT_PARAM")
  SQL.Add("WHERE ROTINAFINFAT = :HROTINAFINFAT")
  SQL.ParamByName("HROTINAFINFAT").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.Active = True

  If SQL.EOF Then
    bsShowMessage("Não foram especificados Grupos/Contratos/Famílias para serem processados", "I")
	Set SQL = Nothing
	Exit Sub
  End If


  'Se for Apropriação de faturamento antecipado
  If (CurrentQuery.FieldByName("TABTIPOPROCESSO").AsInteger = 5) Then
    Set Obj = CreateBennerObject("SAMFaturamento.Apropriacao")
    Obj.ProcessarApropriacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
    Set Obj = Nothing
  Else
    If VisibleMode Then
      Set Obj = CreateBennerObject("BSINTERFACE0016.RotinaFaturamentoBeneficiarios")
      Obj.ProcessarFaturamento(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, True)
    Else
      SQL.Clear
      SQL.Add("SELECT SFAT.DESCRICAO DESCRICAOTIPOFATURAMENTO,")
      SQL.Add("       CFIN.COMPETENCIA,")
      SQL.Add("       RFIN.SEQUENCIA")
      SQL.Add("FROM SFN_ROTINAFIN       RFIN")
      SQL.Add("JOIN SFN_COMPETFIN       CFIN ON RFIN.COMPETFIN       = CFIN.HANDLE")
      SQL.Add("JOIN SIS_TIPOFATURAMENTO SFAT ON CFIN.TIPOFATURAMENTO = SFAT.HANDLE")
      SQL.Add("WHERE RFIN.HANDLE = :HROTINAFIN")
      SQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
      SQL.Active = True

      Dim vsMensagemErro As String
      Dim viRetorno As Long
      Dim vcContainer As CSDContainer
      Set vcContainer = NewContainer

      vcContainer.AddFields("HANDLE:INTEGER;OPCAOCANCELAMENTO:STRING;HOPCAOCANCELAMENTO:INTEGER")
      vcContainer.Insert
      vcContainer.Field("HANDLE").AsInteger             = CurrentQuery.FieldByName("HANDLE").AsInteger
      vcContainer.Field("OPCAOCANCELAMENTO").AsString   = "T"
      vcContainer.Field("HOPCAOCANCELAMENTO").AsInteger = 0

      Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
      viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                       "BSBen018", _
                                       "RotinaFaturamentoBeneficiarios_Processar", _
                                       "Rotina de Faturamento de Beneficiários (Processar) -" + _
                                         " Faturamento: " + SQL.FieldByName("DESCRICAOTIPOFATURAMENTO").AsString + _
                                         " Competência: " + Str(Format(SQL.FieldByName("COMPETENCIA").AsDateTime, "mm/yyyy")) + _
                                         " Sequência: "   + SQL.FieldByName("SEQUENCIA").AsString, _
                                       CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                       "SFN_ROTINAFINFAT", _
                                       "SITUACAOFATURAMENTO", _
                                       "", _
                                       "", _
                                       "P", _
                                       False, _
                                       vsMensagemErro, _
                                       vcContainer)
      If viRetorno = 0 Then
        bsShowMessage("Processo enviado para execução no servidor!", "I")
      Else
        bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
      End If
    End If

    Set Obj = Nothing
    Set SQL = Nothing
  End If

  If VisibleMode Then
    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If
End Sub

Public Sub BOTAOCANCELAR_OnClick()
  If CurrentQuery.State <> 1 Then
	bsShowMessage("Os parâmetros não podem estar em edição", "I")
	Exit Sub
  End If

  Dim Obj As Object

  'Se for Apropriação de faturamento antecipado
  If (CurrentQuery.FieldByName("TABTIPOPROCESSO").AsInteger = 5) Then
    Set Obj = CreateBennerObject("SAMFaturamento.Apropriacao")
    Obj.CancelarApropriacao(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    'SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
    RefreshNodesWithTable("SFN_ROTINAFINFAT")

  Else

    If VisibleMode Then
      Set Obj = CreateBennerObject("BSINTERFACE0016.RotinaFaturamentoBeneficiarios")
      Obj.CancelarFaturamento(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "T", 0)
    Else
      Dim SQL As Object
      Set SQL = NewQuery

      SQL.Clear
      SQL.Add("SELECT SFAT.DESCRICAO DESCRICAOTIPOFATURAMENTO,")
      SQL.Add("       CFIN.COMPETENCIA,")
      SQL.Add("       RFIN.SEQUENCIA")
      SQL.Add("FROM SFN_ROTINAFIN       RFIN")
      SQL.Add("JOIN SFN_COMPETFIN       CFIN ON RFIN.COMPETFIN       = CFIN.HANDLE")
      SQL.Add("JOIN SIS_TIPOFATURAMENTO SFAT ON CFIN.TIPOFATURAMENTO = SFAT.HANDLE")
      SQL.Add("WHERE RFIN.HANDLE = :HROTINAFIN")
      SQL.ParamByName("HROTINAFIN").Value = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
      SQL.Active = True

      Dim vsMensagemErro As String
      Dim viRetorno As Long
      Dim vcContainer As CSDContainer
      Set vcContainer = NewContainer

      vcContainer.AddFields("HANDLE:INTEGER;OPCAOCANCELAMENTO:STRING;HOPCAOCANCELAMENTO:INTEGER")
      vcContainer.Insert
      vcContainer.Field("HANDLE").AsInteger             = CurrentQuery.FieldByName("HANDLE").AsInteger
      vcContainer.Field("OPCAOCANCELAMENTO").AsString   = "T"
      vcContainer.Field("HOPCAOCANCELAMENTO").AsInteger = 0

      Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
      viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                       "BSBen018", _
                                       "RotinaFaturamentoBeneficiarios_Cancelar", _
                                       "Rotina de Faturamento de Beneficiários (Cancelar) -" + _
                                       " Faturamento: " + SQL.FieldByName("DESCRICAOTIPOFATURAMENTO").AsString + _
                                       " Competência: " + Str(Format(SQL.FieldByName("COMPETENCIA").AsDateTime, "mm/yyyy")) + _
                                       " Sequência: "   + SQL.FieldByName("SEQUENCIA").AsString, _
                                       CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                       "SFN_ROTINAFINFAT", _
                                       "SITUACAOFATURAMENTO", _
                                       "SITUACAOPARCELAMENTO", _
                                       "Parcelamento está processado.", _
                                       "C", _
                                       False, _
                                       vsMensagemErro, _
                                       vcContainer)

      Set SQL         = Nothing
      Set vcContainer = Nothing

      If viRetorno = 0 Then
        bsShowMessage("Processo enviado para execução no servidor!", "I")
      Else
        bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
      End If
    End If
  End If

  Set Obj = Nothing

  If VisibleMode Then
    CurrentQuery.Active = False
    CurrentQuery.Active = True
  End If
End Sub

Public Sub BOTAODUPLICAR_OnClick()

  If CurrentQuery.State <> 1 Then
		bsShowMessage("Os parâmetros não podem estar em edição", "I")
		Exit Sub
  End If

  Dim INTERFACE0002 As Object
  Dim vsMensagem As String


  Set INTERFACE0002 = CreateBennerObject("BSINTERFACE0002.GerarFormularioVirtual")

	INTERFACE0002.Exec(CurrentSystem, _
					   1, _
					   "TV_FORM0102", _
					   "Duplicar rotina financeira",  _
					   0, _
					   320, _
					   530, _
					   False, _
					   vsMensagem, _
					   Null)


  Set INTERFACE0002 = Nothing

  WriteAudit("D", HandleOfTable("SFN_ROTINAFINFAT"), CurrentQuery.FieldByName("HANDLE").AsInteger,"Faturamento de Beneficiários- Duplicação")
End Sub

Public Sub VerificaSeProcessada(CanContinue As Boolean)
  If CurrentQuery.FieldByName("SITUACAOFATURAMENTO").Value <> "1" Then
	CanContinue = False
	bsShowMessage("A Rotina não está aberta", "E")
	Exit Sub
  End If
End Sub

Public Sub BOTAOPROCESSARPARCELAR_OnClick()
  Dim BSINTERFACE As Object
  Dim BSSERVEREXEC As Object
  Dim vsMensagemErro As String
  Dim viRetorno As Long

  If Not (CurrentQuery.State = 1) Then
	bsShowMessage("O registro não pode estar em edição", "I")
	Exit Sub
  End If

  If (VisibleMode) Then
    Set BSINTERFACE = CreateBennerObject("BSINTERFACE0016.RotinaFaturamentoBeneficiarios")
    BSINTERFACE.ProcessarParcelarFaturamento(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, True)
  Else
    Set BSSERVEREXEC = CreateBennerObject("BSSERVEREXEC.ProcessosServidor")
    viRetorno = BSSERVEREXEC.ExecucaoImediata(CurrentSystem, _
	                        		          "BSBEN018", _
	                                		  "RotinaFaturamentoParcelamentoBeneficiarios_Processar", _
			                                  "Processamento de Faturamento e Parcelamento da Rotina Financeira: " + _
			                                  CurrentQuery.FieldByName("HANDLE").AsString, _
			                                  CurrentQuery.FieldByName("HANDLE").AsInteger, _
			                                  "SFN_ROTINAFINFAT", _
			                                  "SITUACAOFATURAMENTO", _
			                                  "", _
			                                  "", _
			                                  "P", _
			                                  False, _
			                                  vsMensagemErro, _
			                                  Null)
    If (viRetorno = 0) Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If
  End If

  Set BSINTERFACE = Nothing
  Set BSSERVEREXEC = Nothing
End Sub

Public Sub BOTAORH_OnClick()
  Dim Obj As Object
  Dim SQL As Object
  Dim vCompetencia As Date
  Dim vDataRotina As Date
  Dim vTipoFaturamento As Integer
  Dim vRotinaFinFat As Integer

  If CurrentQuery.State <> 1 Then
		bsShowMessage("Os parâmetros não podem estar em edição", "I")
		Exit Sub
  End If

  Set SQL = NewQuery

  SQL.Add("SELECT A.COMPETENCIA COMPETENCIA, B.DESCRICAO, B.DATAROTINA DATAROTINA, B.SEQUENCIA SEQROTINA, D.HANDLE TIPOFATURAMENTO")
  SQL.Add("FROM SFN_COMPETFIN A, SFN_ROTINAFIN B, SFN_ROTINAFINFAT C, SIS_TIPOFATURAMENTO D")
  SQL.Add("WHERE B.HANDLE =:ROTFINFAT AND B.COMPETFIN = A.HANDLE AND A.TIPOFATURAMENTO = D.HANDLE")

  SQL.ParamByName("ROTFINFAT").Value = CurrentQuery.FieldByName("ROTINAFIN").Value
  SQL.Active = True

  vCompetencia = SQL.FieldByName("COMPETENCIA").AsDateTime
  vDataRotina = SQL.FieldByName("DATAROTINA").AsDateTime
  vTipoFaturamento = SQL.FieldByName("TIPOFATURAMENTO").AsInteger
  vRotinaFinFat = CurrentQuery.FieldByName("HANDLE").Value

  If VisibleMode Then
    Set Obj = CreateBennerObject("RotArq.Rotinas")
    Obj.ArquivoRh(CurrentSystem, vCompetencia, vDataRotina, vTipoFaturamento, vRotinaFinFat, 0)
    Set Obj = Nothing
  Else
    Dim vsMensagemErro As String
    Dim viRetorno As Long
    Dim vcContainer As CSDContainer

    Set vcContainer = NewContainer
    vcContainer.AddFields("COMPETENCIA:TDATETIME;DATAROTINA:TDATETIME;TIPOFATURAMENTO:INTEGER;ROTINAFINFAT:INTEGER;CONTRATO:INTEGER")
    vcContainer.Insert
    vcContainer.Field("COMPETENCIA").AsDateTime = vCompetencia
    vcContainer.Field("DATAROTINA").AsDateTime = vDataRotina
    vcContainer.Field("TIPOFATURAMENTO").AsInteger = vTipoFaturamento
    vcContainer.Field("ROTINAFINFAT").AsInteger = vRotinaFinFat
    vcContainer.Field("CONTRATO").AsInteger = 0

    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
	                                "RotArq", _
	                                "ArquivoRH_Exec", _
	                                "Geração de Arquivo para RH: " + _
	                                SQL.FieldByName("COMPETENCIA").AsString + " - " + SQL.FieldByName("SEQROTINA").AsString + _
									SQL.FieldByName("DESCRICAO").AsString, _
	                                0, _
	                                "", _
	                                "", _
	                                "", _
	                                "", _
	                                "P", _
	                                False, _
	                                vsMensagemErro, _
	                                vcContainer)

    If viRetorno = 0 Then
      bsShowMessage("Processo enviado para execução no servidor!", "I")
    Else
      bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
    End If
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  'Luciano T. Alberti - SMS 91413 - 24/01/2008 - Início
  Dim qRotinaFin As Object
  Set qRotinaFin = NewQuery
  With qRotinaFin
    .Active = False
    .Clear
    .Add("SELECT SITUACAO")
    .Add("  FROM SFN_ROTINAFIN")
    .Add(" WHERE HANDLE = :HANDLE")
    .ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("ROTINAFIN").AsInteger
    .Active = True
  End With
  Set qRotinaFin = Nothing
  'Luciano T. Alberti - SMS 91413 - 24/01/2008 - Fim

  Dim SQL As Object

  Set SQL = NewQuery

  'SMS 30184 - Cazangi - Início
  SQL.Add("SELECT CODIGO")
  SQL.Add("FROM SIS_TIPOFATURAMENTO")
  SQL.Add("WHERE HANDLE = (SELECT TIPOFATURAMENTO FROM SFN_ROTINAFIN WHERE HANDLE = :HANDLE)")

  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SFN_ROTINAFIN")
  SQL.Active = True
  'SMS 30184 - Cazangi - Fim

  If SQL.FieldByName("CODIGO").AsInteger = 110 Then
		CALCULAPF.Caption                 = "Serviço"
		CALCULACONTRIBUICAO.Visible       = False
		SEQUENCIASALARIO.Visible          = False
		CALCULAMODULOCOBERTURAPF.Visible  = False
		BOTAOPARCELAR.Visible             = False
		BOTAOCANCELARPARCELAMENTO.Visible = False
  ElseIf SQL.FieldByName("CODIGO").AsInteger = 120 Then
		CALCULAPF.Caption                 = "PF"
		CALCULACONTRIBUICAO.Visible       = False
		SEQUENCIASALARIO.Visible          = False
		CALCULAMODULOCOBERTURAPF.Visible  = False
		BOTAOPARCELAR.Visible             = False
		BOTAOCANCELARPARCELAMENTO.Visible = False
  ElseIf SQL.FieldByName("CODIGO").AsInteger = 140 Then
		CALCULAPF.Caption                 = "PF"
		CALCULACONTRIBUICAO.Visible       = False
		SEQUENCIASALARIO.Visible          = False
		CALCULAMODULOCOBERTURAPF.Visible  = False
		BOTAOPARCELAR.Visible             = False
		BOTAOCANCELARPARCELAMENTO.Visible = False
  Else
		CALCULAPF.Caption = "PF e parcelamento"
		CALCULACONTRIBUICAO.Visible       = True
		SEQUENCIASALARIO.Visible          = True
		CALCULAMODULOCOBERTURAPF.Visible  = True
        If (CurrentQuery.FieldByName("SITUACAOPARCELAMENTO").AsString = "1") And _
           (CurrentQuery.FieldByName("SITUACAOFATURAMENTO").AsString = "5") Then
	      BOTAOPARCELAR.Visible             = True
	    Else
          BOTAOPARCELAR.Visible             = False
        End If
		If CurrentQuery.FieldByName("SITUACAOPARCELAMENTO").AsString = "5" Then
          BOTAOCANCELARPARCELAMENTO.Visible = True
        Else
          BOTAOCANCELARPARCELAMENTO.Visible = False
        End If
  End If

  If CurrentQuery.FieldByName("SITUACAOFATURAMENTO").AsString = "1" Then
	BOTAOPROCESSAR.Visible             = True
  Else
    BOTAOPROCESSAR.Visible             = False
  End If

  If CurrentQuery.FieldByName("SITUACAOFATURAMENTO").AsString = "5" Then
	BOTAOCANCELAR.Visible             = True
  Else
    BOTAOCANCELAR.Visible             = False
  End If

  Set SQL = Nothing
  Set SQL2 = Nothing

  If (VisibleMode) Then
	TESOURARIA.ReadOnly = Not (CurrentQuery.FieldByName("TABTIPOPROCESSO").AsInteger = 3)
	TIPODOCUMENTO.ReadOnly = Not (CurrentQuery.FieldByName("TABTIPOPROCESSO").AsInteger = 3)
  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "BOTAOPARCELAR"
			BOTAOPARCELAR_OnClick
		Case "BOTAOCANCELARPARCELAMENTO"
			BOTAOCANCELARPARCELAMENTO_OnClick
		Case "BOTAODUPLICAR"
			BOTAODUPLICAR_OnClick
		Case "BOTAORH"
			BOTAORH_OnClick
		Case "BOTAOPROCESSARPARCELAR"
			BOTAOPROCESSARPARCELAR_OnClick
	End Select
End Sub

Public Sub TABTIPOPROCESSO_OnChange()
  If (TABTIPOPROCESSO.PageIndex = 1) Then
	TESOURARIA.ReadOnly = False
	TIPODOCUMENTO.ReadOnly = False
  Else
	TIPODOCUMENTO.ReadOnly = True
	TESOURARIA.ReadOnly = True

	CurrentQuery.FieldByName("TESOURARIA").Clear
	CurrentQuery.FieldByName("TIPODOCUMENTO").Clear
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  VerificaSeProcessada(CanContinue)
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  VerificaSeProcessada(CanContinue)

  Dim qRotinaFin As Object
  Set qRotinaFin = NewQuery

  qRotinaFin.Clear
  qRotinaFin.Add("SELECT HANDLE FROM SFN_ROTINAFINFAT_PARAM WHERE ROTINAFINFAT =:PROTINA")
  qRotinaFin.ParamByName("PROTINA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qRotinaFin.Active = True

  If qRotinaFin.FieldByName("HANDLE").AsInteger > 0 Then
    bsShowMessage("Excluir antes os itens do subitem 'Contratos a faturar'!", "E")
    Set qRotinaFin = Nothing
    CanContinue = False
    Exit Sub
  End If

  Set qRotinaFin = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)


	If ((CurrentQuery.FieldByName("TABTIPOPROCESSO").AsInteger = 2) Or (CurrentQuery.FieldByName("TABTIPOPROCESSO").AsInteger = 3)) And _
       (CurrentQuery.FieldByName("LOCALFATURAMENTO").IsNull) Then
   		bsShowMessage("Informar o local de faturamento!", "E")
   		CanContinue = False
   		Exit Sub
	End If



	    If ((CurrentQuery.FieldByName("TABTIPOPROCESSO").AsInteger = 1) Or (CurrentQuery.FieldByName("TABTIPOPROCESSO").AsInteger = 4)) Then
	   	   CurrentQuery.FieldByName("LOCALFATURAMENTO").AsString = "C"
	    End If

		Dim SQLCompetFin As Object
		Dim qEmpresa As Object 'sms 42272 - Edilson.Castro - 18-05-2005
		Set SQLCompetFin = NewQuery

		SQLCompetFin.Add("SELECT c.CODIGO FROM SFN_ROTINAFIN A, SFN_COMPETFIN B, SIS_TIPOFATURAMENTO C")
		SQLCompetFin.Add("WHERE A.HANDLE = :HANDLEA")
		SQLCompetFin.Add("AND B.HANDLE = A.COMPETFIN")
		SQLCompetFin.Add("AND C.HANDLE = B.TIPOFATURAMENTO")

		SQLCompetFin.ParamByName("HANDLEA").Value = CurrentQuery.FieldByName("ROTINAFIN").Value
		SQLCompetFin.Active = True

		If (CurrentQuery.FieldByName("TABTIPOPROCESSO").Value = 2) Then
			If (SQLCompetFin.FieldByName("CODIGO").Value <> 110) Then '  <> Custo Operacional
				If (CurrentQuery.FieldByName("CALCULAMENSALIDADE").AsString = "S") And _
					 (CurrentQuery.FieldByName("CALCULAPF").AsString = "N") Then
					If WebMode Then
                      bsShowMessage("Foi marcado para calcular Mensalidade e não foi marcado para calcular PF. Para que seja calculado a PF deverá ser marcado o campo 'PF e Parcelamento'", "I")
					Else
					  If bsShowMessage("Foi marcado para calcular Mensalidade e não foi marcado para calcular PF. Deseja calcular PF?", "Q") = vbYes Then
						CurrentQuery.FieldByName("CALCULAPF").Value = "S"
					  End If
					End If
				End If
			End If
		End If

		' Se <> autogestao e cota patronal da erro
		If (SQLCompetFin.FieldByName("CODIGO").Value <> 130) Then
			If (CurrentQuery.FieldByName("TABTIPOPROCESSO").Value = 1) Then
				CanContinue = False
				bsShowMessage("Cota patronal permitida somente para AutoGestão", "E")
			ElseIf (CurrentQuery.FieldByName("TABTIPOPROCESSO").Value = 4) Then
				CanContinue = False
				bsShowMessage("Taxa Admnistração permitida somente para AutoGestão", "E")
			ElseIf (CurrentQuery.FieldByName("CALCULACONTRIBUICAO").Value = "S") Then
				CanContinue = False
				bsShowMessage("Contribuição social permitida somente para AutoGestão", "E")
			ElseIf (CurrentQuery.FieldByName("CALCULAMODULOCOBERTURAPF").Value = "S") Then
				CanContinue = False
				bsShowMessage("Módulo cobertura de PF permitido somente para AutoGestão", "E")
			End If

			If (Not CurrentQuery.FieldByName("SEQUENCIASALARIO").IsNull) Then
				CanContinue = False
				bsShowMessage("Informar a sequência do salário somente para AutoGestão", "E")
			End If
		End If

		' Se Custo Operacional - Rateio
		If VerificaSeRateio Then
			If CurrentQuery.FieldByName("TABTIPOPROCESSO").Value = 1 Then
				CanContinue = False
				bsShowMessage("Contribuição não permitida para Custo Operacional - Rateio", "E")
			End If

			If CurrentQuery.FieldByName("TABTIPOPROCESSO").Value = 2 And _
				CurrentQuery.FieldByName("CALCULACONTRIBUICAO").Value = "S" Then
				CanContinue = False
				bsShowMessage("Calcula contribuição não permitida para Custo Operacional - Rateio", "E")
			End If

			If CurrentQuery.FieldByName("TABTIPOPROCESSO").Value = 3 Then
				If CurrentQuery.FieldByName("INSCRICAOFATURARATE").Value <> "0" Then
					CanContinue = False
					bsShowMessage("Custo Operacional - Rateio só permite faturar até inscrição", "E")
				End If
			End If
		End If

		If CanContinue = False Then
			SQLCompetFin.Active = False
			Set SQLCompetFin = Nothing
			Exit Sub
		End If

		If (CurrentQuery.FieldByName("TABFILTROVENCIMENTO").AsInteger = 2) Then ' sem filtro
			CurrentQuery.FieldByName("DIAVENCIMENTOINICIAL").Value = Null
			CurrentQuery.FieldByName("DIAVENCIMENTOFINAL").Value = Null
			CurrentQuery.FieldByName("VENCIMENTONOMESSEGUINTE").Value = "N"
		Else
			If (CurrentQuery.FieldByName("DIAVENCIMENTOINICIAL").AsInteger > _
				 CurrentQuery.FieldByName("DIAVENCIMENTOFINAL").AsInteger) And _
				 CurrentQuery.FieldByName("VENCIMENTONOMESSEGUINTE").AsString = "N" Then
				CanContinue = False
				bsShowMessage("Dia de vencimento final só pode ser menor que inicial para vencimento no mês seguinte", "E")
				Exit Sub
			End If
		End If

		If (CurrentQuery.State = 3) And _
			 ((CurrentQuery.FieldByName("TABTIPOPROCESSO").Value = 1) Or _
			  (CurrentQuery.FieldByName("TABTIPOPROCESSO").Value = 4)) Then
			CurrentQuery.FieldByName("LOCALFATURAMENTO").Value = "C"
		End If

		If (CurrentQuery.FieldByName("TABTIPOPROCESSO").Value = 1) Then ' Cota patronal
			CurrentQuery.FieldByName("CALCULACONTRIBUICAO").Value = "N"
			CurrentQuery.FieldByName("CALCULAMENSALIDADE").Value = "N"
			CurrentQuery.FieldByName("CALCULAPF").Value = "N"
			CurrentQuery.FieldByName("CALCULAMODULOCOBERTURAPF").Value = "N"
		Else
			If (CurrentQuery.FieldByName("TABTIPOPROCESSO").Value = 2) Then ' Mensal
				CurrentQuery.FieldByName("INSCRICAOFATURARATE").Value = "2"

				If CurrentQuery.FieldByName("CALCULAMENSALIDADE").Value = "S" Or _
					 CurrentQuery.FieldByName("CALCULACONTRIBUICAO").Value = "S" Or _
					 CurrentQuery.FieldByName("CALCULAPF").Value = "S" Then
					CurrentQuery.FieldByName("CALCULAMODULOCOBERTURAPF").Value = "N"
				End If

				If CurrentQuery.FieldByName("CALCULAMODULOCOBERTURAPF").Value = "S" Then
					CurrentQuery.FieldByName("CALCULACONTRIBUICAO").Value = "N"
					CurrentQuery.FieldByName("CALCULAMENSALIDADE").Value = "N"
					CurrentQuery.FieldByName("CALCULAPF").Value = "N"
				End If
			Else
				If (CurrentQuery.FieldByName("TABTIPOPROCESSO").Value = 3) Then ' Inscricao
					CurrentQuery.FieldByName("CALCULAPF").Value = "N"
					CurrentQuery.FieldByName("CALCULAMODULOCOBERTURAPF").Value = "N"
				End If
			End If
		End If

		If (CurrentQuery.FieldByName("TABTIPOPROCESSO").Value = 2) And _
			 (CurrentQuery.FieldByName("CALCULACONTRIBUICAO").Value = "N") And _
			 (CurrentQuery.FieldByName("CALCULAMENSALIDADE").Value = "N") And _
			 (CurrentQuery.FieldByName("CALCULAPF").Value = "N") Then ' Mensal
			If SQLCompetFin.FieldByName("CODIGO").Value = 130 Then
				If CurrentQuery.FieldByName("CALCULAMODULOCOBERTURAPF").Value = "N" Then
					CanContinue = False
					bsShowMessage("Deve selecionar Contribuição, Mensalidade, PF ou Módulo Cobertura de PF", "E")
				End If
			Else
				CanContinue = False
				bsShowMessage("Deve selecionar Mensalidade e/ou PF", "E")
			End If
		Else
			If (CurrentQuery.FieldByName("TABTIPOPROCESSO").Value = 3) And _
				 (CurrentQuery.FieldByName("CALCULACONTRIBUICAO").Value = "N") And _
				 (CurrentQuery.FieldByName("CALCULAMENSALIDADE").Value = "N") Then ' Inscricao
				CanContinue = False
				bsShowMessage("Deve selecionar Contribuição e/ou Mensalidade", "E")
			End If
		End If

		SQLCompetFin.Active = False

		Set SQLCompetFin = Nothing
End Sub

Public Function VerificaSeRateio As Boolean
	Dim SQLRotFin As Object
	Set SQLRotFin = NewQuery

	SQLRotFin.Add("SELECT C.CODIGO FROM SFN_ROTINAFIN A, SFN_COMPETFIN B, SIS_TIPOFATURAMENTO C")
	SQLRotFin.Add("WHERE A.HANDLE = :HANDLE")
	SQLRotFin.Add("  AND B.HANDLE = A.COMPETFIN")
	SQLRotFin.Add("  AND C.HANDLE = B.TIPOFATURAMENTO")

	SQLRotFin.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ROTINAFIN").Value
	SQLRotFin.Active = True

	If SQLRotFin.FieldByName("CODIGO").Value = 140 Then
		VerificaSeRateio = True
	Else
		VerificaSeRateio = False
	End If

	SQLRotFin.Active = False

	Set SQLRotFin = Nothing
End Function
