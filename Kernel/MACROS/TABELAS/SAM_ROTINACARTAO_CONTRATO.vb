'HASH: A4A72F689F3A42AAE15C6A194A364601
 
'Macro: SAM_ROTINACARTAO_CONTRATO
'A funcao NodeInternalCode é utilizada para determinar se a carga correspondente é da Tarefas de Modelo,
'sendo, mostra o Tab - Modelo para agendamento, não sendo, mostra o Tab - Rotina
'Alteração: 26/12/2005
'      SMS: 52120 - Marcelo Barbosa
'#Uses "*bsShowMessage"

Public Sub BOTAOCANCELAR_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <> 1 Then
	bsShowMessage("Os parâmetros não podem estar em edição.", "I")
	Exit Sub
  End If

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT USUARIOGERACAO, DATACANCELAR, MOTIVOCANCELAR")
  SQL.Add("FROM SAM_ROTINACARTAO")
  SQL.Add("WHERE HANDLE = :HROTINACARTAO")

  SQL.ParamByName("HROTINACARTAO").Value = CurrentQuery.FieldByName("ROTINACARTAO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("USUARIOGERACAO").IsNull Then
    bsShowMessage("A Geração ainda não foi processada.", "I")
    Set SQL = Nothing
    Exit Sub
  ElseIf SQL.FieldByName("DATACANCELAR").IsNull Or SQL.FieldByName("MOTIVOCANCELAR").IsNull Then
    bsShowmessage("Faltam parâmetros de cancelamento na Rotina Cartão.", "I")
    Set SQL = Nothing
    Exit Sub
  End If

  Set SQL = Nothing

  If bsShowMessage("Confirma o cancelamento dos cartões?", "Q") = vbYes Then

    If VisibleMode Then
    	Set Obj = CreateBennerObject("BSInterface0004.RotinaCartao")
    	Obj.CancelarContrato(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Else
    	Dim vsMensagemErro As String
   		Dim viRetorno As Long
   	    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
   	    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBEN009", _
                                     "RotinaCartao_CancelarContrato", _
                                     "Rotina de Cancelamento de Cartão - Rotina: " + CStr(CurrentQuery.FieldByName("ROTINACARTAO").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_ROTINACARTAO_CONTRATO", _
                                     "SITUACAOCANCELAMENTO", _
                                     "", _
                                     "", _
                                     "P", _
                                     True, _
                                     vsMensagemErro, _
                                     Null)

	   	If viRetorno = 0 Then
     		bsShowMessage("Processo enviado para execução no servidor!", "I")
   		Else
     		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
   		End If
    End If

    Set Obj = Nothing
  End If
End Sub

Public Sub BOTAODESBLOQUEAR_OnClick()
  Dim Obj As Object

  If CurrentQuery.State <> 1 Then
    bsShowMessage("Os parâmetros não podem estar em edição.", "I")
    Exit Sub
  End If

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT USUARIOGERACAO")
  SQL.Add("FROM SAM_ROTINACARTAO")
  SQL.Add("WHERE HANDLE = :HROTINACARTAO")

  SQL.ParamByName("HROTINACARTAO").Value = CurrentQuery.FieldByName("ROTINACARTAO").AsInteger
  SQL.Active = True

  If SQL.FieldByName("USUARIOGERACAO").IsNull Then
    bsShowMessage("A Geração ainda não foi processada.", "I")
    Set SQL = Nothing
    Exit Sub
  End If

  Set SQL = Nothing

  If bsShowMessage("Confirma o desbloqueio dos cartões?", "Q") = vbYes Then
	If VisibleMode Then
	    Set Obj = CreateBennerObject("BSInterface0004.RotinaCartao")
    	Obj.DesbloquearContrato(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
    Else
	    Dim vsMensagemErro As String
   		Dim viRetorno As Long
   	    Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
   	    viRetorno = Obj.ExecucaoImediata(CurrentSystem, _
                                     "BSBEN009", _
                                     "RotinaCartao_DesbloquearContrato", _
                                     "Rotina de Desbloqueio de Cartão - Rotina: " + CStr(CurrentQuery.FieldByName("ROTINACARTAO").AsInteger), _
                                     CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                     "SAM_ROTINACARTAO_CONTRATO", _
                                     "SITUACAODESBLOQUEIO", _
                                     "", _
                                     "", _
                                     "P", _
                                     True, _
                                      vsMensagemErro, _
                                     Null)


	   	If viRetorno = 0 Then
     		bsShowMessage("Processo enviado para execução no servidor!", "I")
   		Else
     		bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
   		End If
	End If

    Set Obj = Nothing
  End If
End Sub

Public Sub CARTAOMOTIVOEMISSAO_OnPopup(ShowPopup As Boolean)
  UpdateLastUpdate("SAM_CONTRATO_CARTAOMOTIVO")
End Sub

Public Sub TABLE_AfterScroll()

  If WebMode Then
  	CARTAOMOTIVOEMISSAO.WebLocalWhere = "A.HANDLE IN (SELECT CARTAOMOTIVO  " + _
            						    "FROM SAM_CONTRATO_CARTAOMOTIVO    " + _
               							"WHERE CONTRATO = @CAMPO(CONTRATO))"
  ElseIf VisibleMode Then
    CARTAOMOTIVOEMISSAO.LocalWhere = "HANDLE IN (SELECT CARTAOMOTIVO  " + _
            						 "FROM SAM_CONTRATO_CARTAOMOTIVO    " + _
               						 "WHERE CONTRATO = @CONTRATO)"
  End If




  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT USUARIOGERACAO, tabtipogeracao ")
  SQL.Add("  FROM SAM_ROTINACARTAO")
  SQL.Add(" WHERE HANDLE = :HROTINACARTAO")

  SQL.ParamByName("HROTINACARTAO").Value = RecordHandleOfTable("SAM_ROTINACARTAO")
  SQL.Active = True

  'SMS 52120 - Marcelo Barbosa - 26/12/2005
  If VisibleMode Then
	If NodeInternalCode <> 502 Then 'Verifica qual a carga para habilitar as opções correspondentes (Rotina ou Modelo para agendamento)
	  If SQL.FieldByName("tabtipogeracao").Value = 3 Then
		BOTAOCANCELAR.Visible = False
		BOTAODESBLOQUEAR.Visible = False
		CARTAOMOTIVOEMISSAO.Visible = False
		CORRESPONDENCIAANO.Visible = False
		CORRESPONDENCIANUMERO.Visible = False
	  Else
		BOTAOCANCELAR.Visible = True
		BOTAODESBLOQUEAR.Visible = True
		CARTAOMOTIVOEMISSAO.Visible = True
		CORRESPONDENCIAANO.Visible = True
		CORRESPONDENCIANUMERO.Visible = True
	  End If
	Else
	  BOTAOCANCELAR.Visible = False
	  BOTAODESBLOQUEAR.Visible = False

	  If SQL.FieldByName("tabtipogeracao").Value = 3 Then
		CARTAOMOTIVOEMISSAO.Visible = False
		CORRESPONDENCIAANO.Visible = False
		CORRESPONDENCIANUMERO.Visible = False
	  Else
		CARTAOMOTIVOEMISSAO.Visible = True
		CORRESPONDENCIAANO.Visible = True
		CORRESPONDENCIANUMERO.Visible = True
	  End If
    End If
  End If

  Set SQL = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  CanContinue =False

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT USUARIOGERACAO, tabtipogeracao ")
  SQL.Add("  FROM SAM_ROTINACARTAO")
  SQL.Add(" WHERE HANDLE = :HROTINACARTAO")

  SQL.ParamByName("HROTINACARTAO").Value = RecordHandleOfTable("SAM_ROTINACARTAO")
  SQL.Active = True

  'SMS 52120 - Marcelo Barbosa - 26/12/2005
  If VisibleMode Or WebMode Then
	If NodeInternalCode <> 502 Then 'Verifica qual a carga para habilitar as opções correspondentes (Rotina ou Modelo para agendamento)
	  If SQL.FieldByName("USUARIOGERACAO").IsNull Then
	  Else
		bsShowMessage("Geração já foi processada.", "E")
		Set SQL = Nothing
		CanContinue = False
		Exit Sub
	  End If

	  If (SQL.FieldByName("tabtipogeracao").Value = 4) Then
		bsShowMessage("Tipo de geração não aceita contrato.", "E")
		Set SQL = Nothing
		CanContinue = False
		Exit Sub
	  End If

	  If SQL.FieldByName("tabtipogeracao").Value = 3 Then
		BOTAOCANCELAR.Visible = False
		BOTAODESBLOQUEAR.Visible = False
		CARTAOMOTIVOEMISSAO.Visible = False
		CORRESPONDENCIAANO.Visible = False
		CORRESPONDENCIANUMERO.Visible = False
	  Else
		BOTAOCANCELAR.Visible = True
		BOTAODESBLOQUEAR.Visible = True
		CARTAOMOTIVOEMISSAO.Visible = True
		CORRESPONDENCIAANO.Visible = True
		CORRESPONDENCIANUMERO.Visible = True
	  End If
	Else
	  If(SQL.FieldByName("tabtipogeracao").Value <> 1) And (SQL.FieldByName("tabtipogeracao").Value <> 3)Then
		bsShowMessage("Tipo de geração não aceita contrato.", "E")
		Set SQL = Nothing
		CanContinue = False
		Exit Sub
	  End If

	  If SQL.FieldByName("tabtipogeracao").Value = 3 Then
		CARTAOMOTIVOEMISSAO.Visible = False
		CORRESPONDENCIAANO.Visible = False
		CORRESPONDENCIANUMERO.Visible = False
	  Else
		CARTAOMOTIVOEMISSAO.Visible = True
		CORRESPONDENCIAANO.Visible = True
		CORRESPONDENCIANUMERO.Visible = True
	  End If
	End If
  End If

  Set SQL = Nothing

  CanContinue = True
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  CanContinue = False

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT USUARIOGERACAO, tabtipogeracao, tabtipogeracaomodelo")
  SQL.Add("  FROM SAM_ROTINACARTAO")
  SQL.Add(" WHERE HANDLE = :HROTINACARTAO")

  SQL.ParamByName("HROTINACARTAO").Value = RecordHandleOfTable("SAM_ROTINACARTAO")
  SQL.Active = True

  'SMS 52120 - Marcelo Barbosa - 26/12/2005
  If VisibleMode Or WebMode Then
	If NodeInternalCode <> 502 Then 'Verifica qual a carga para habilitar as opções correspondentes (Rotina ou Modelo para agendamento)
	  If CurrentQuery.State = 2 Then
		Dim S As Object
		Set S = NewQuery

		S.Add("SELECT USUARIOGERACAO ")
		S.Add("FROM SAM_ROTINACARTAO ")
		S.Add("WHERE HANDLE = :HANDLE")

		S.ParamByName("HANDLE").Value =CurrentQuery.FieldByName("ROTINACARTAO").AsInteger
		S.Active = True

		If S.FieldByName("USUARIOGERACAO").AsInteger > 0 Then
		  bsShowMessage("Geração já processado, alteração não permitida.", "E")
		  Exit Sub
		End If
	  End If

	  '58165 - Éverton
	  If SQL.FieldByName("tabtipogeracao").Value = 1 Then
		If CurrentQuery.FieldByName("CARTAOMOTIVOEMISSAO").IsNull Then
		  bsShowMessage("O campo 'Motivo emissão' não pode ser nulo!", "E")
		  CanContinue = False
		  Exit Sub
		End If
	  End If
	Else
	  '58165 - Éverton
	  If SQL.FieldByName("tabtipogeracaomodelo").Value = 1 Then
		If CurrentQuery.FieldByName("CARTAOMOTIVOEMISSAO").IsNull Then
		  bsShowMessage("O campo 'Motivo emissão' não pode ser nulo!", "E")
		  CanContinue = False
		  Exit Sub
		End If
	  End If
	End If
  End If

  'sms 48610  verifica se intervalo cadastrado é valido
  Dim qContratoIni As Object
  Set qContratoIni = NewQuery
  Dim qContratoFim As Object
  Set qContratoFim = NewQuery
  Dim qSelFiltro As Object
  Set qSelFiltro = NewQuery

  qContratoIni.Clear

  qContratoIni.Add("SELECT C.CONTRATO  ")
  qContratoIni.Add("  FROM SAM_CONTRATO C")
  qContratoIni.Add(" WHERE C.HANDLE = :HCONTRATOINI ")

  qContratoIni.ParamByName("HCONTRATOINI").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  qContratoIni.Active = True

  qContratoFim.Clear

  qContratoFim.Add("SELECT C.CONTRATO  ")
  qContratoFim.Add("  FROM SAM_CONTRATO C")
  qContratoFim.Add(" WHERE C.HANDLE = :HCONTRATOFIM ")

  qContratoFim.ParamByName("HCONTRATOFIM").AsInteger = CurrentQuery.FieldByName("CONTRATOF").AsInteger
  qContratoFim.Active = True

  If qContratoIni.FieldByName("CONTRATO").AsInteger > qContratoFim.FieldByName("CONTRATO").AsInteger Then
    bsShowMessage("Intervalo de contrato inválido. Contrato inicial maior que contrato final.", "E")
    CanContinue = False
    Exit Sub
  End If

  qSelFiltro.Clear

  qSelFiltro.Add("Select R.HANDLE")
  qSelFiltro.Add("  FROM SAM_ROTINACARTAO_CONTRATO R, SAM_CONTRATO C, SAM_CONTRATO CF   ")
  qSelFiltro.Add(" WHERE C.HANDLE = R.CONTRATO AND CF.HANDLE = R.CONTRATOF AND R.ROTINACARTAO = :ROTINACARTAO ")
  qSelFiltro.Add("       And ((:CONTRATOINI BETWEEN C.CONTRATO And CF.CONTRATO) Or (:CONTRATOFIM BETWEEN C.CONTRATO And CF.CONTRATO)) ")
  qSelFiltro.Add("       AND R.HANDLE <> :HANDLE ")

  qSelFiltro.ParamByName("ROTINACARTAO").AsInteger = CurrentQuery.FieldByName("ROTINACARTAO").AsInteger
  qSelFiltro.ParamByName("CONTRATOINI").AsInteger = qContratoIni.FieldByName("CONTRATO").AsInteger
  qSelFiltro.ParamByName("CONTRATOFIM").AsInteger = qContratoFim.FieldByName("CONTRATO").AsInteger
  qSelFiltro.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qSelFiltro.Active = True

  If Not qSelFiltro.FieldByName("HANDLE").IsNull Then
    bsShowMessage("Intervalo de contrato inválido", "E")
    Exit Sub
  End If

  Set qContratoIni = Nothing
  Set qContratoFim = Nothing
  Set qSelFiltro = Nothing
  'sms 48610

  'Incluído na SMS 51565 - 06.03.2006
  vMsg = VerificaSequencial

  If (vMsg <> "") Then
    bsShowMessage(vMsg, "I")
    Exit Sub
  End If
  'Final SMS 51565

  CanContinue = True
End Sub

Public Sub TABLE_BeforeScroll()
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT USUARIOGERACAO, tabtipogeracao ")
  SQL.Add("  FROM SAM_ROTINACARTAO")
  SQL.Add(" WHERE HANDLE = :HROTINACARTAO")

  SQL.ParamByName("HROTINACARTAO").Value = RecordHandleOfTable("SAM_ROTINACARTAO")
  SQL.Active = True

  'SMS 52120 - Marcelo Barbosa - 26/12/2005
  If VisibleMode Then
	If NodeInternalCode <> 502 Then 'Verifica qual a carga para habilitar as opções correspondentes (Rotina ou Modelo para agendamento)
	  If SQL.FieldByName("tabtipogeracao").Value = 3 Then
		BOTAOCANCELAR.Visible = False
		BOTAODESBLOQUEAR.Visible = False
		CARTAOMOTIVOEMISSAO.Visible = False
		CORRESPONDENCIAANO.Visible = False
		CORRESPONDENCIANUMERO.Visible = False
	  Else
		BOTAOCANCELAR.Visible = True
		BOTAODESBLOQUEAR.Visible = True
		CARTAOMOTIVOEMISSAO.Visible = True
		CORRESPONDENCIAANO.Visible = True
		CORRESPONDENCIANUMERO.Visible = True
	  End If
	Else
	  BOTAOCANCELAR.Visible = False
	  BOTAODESBLOQUEAR.Visible = False

	  If SQL.FieldByName("tabtipogeracao").Value = 3 Then
		CARTAOMOTIVOEMISSAO.Visible = False
		CORRESPONDENCIAANO.Visible = False
		CORRESPONDENCIANUMERO.Visible = False
	  Else
		CARTAOMOTIVOEMISSAO.Visible = True
		CORRESPONDENCIAANO.Visible = True
		CORRESPONDENCIANUMERO.Visible = True
	  End If
	End If
  End If

  Set SQL = Nothing
End Sub

' Criado na SMS 51565 - 06.03.2006
Public Function VerificaSequencial As String
  Dim qSequencia As Object
  Dim qRotina As Object
  Dim qContratoIni As Object
  Dim qContratoFim As Object
  Dim vMsg As String
  Dim vsequencial As Long
  Dim vSequenciaAtual As Long

  VerificaSequencial = ""

  Set qSequencia = NewQuery

  qSequencia.Clear

  qSequencia.Add("SELECT CONTROLESEQUENCIACONTRATO ")
  qSequencia.Add("  FROM SAM_PARAMETROSBENEFICIARIO ")

  qSequencia.Active = True

  If (qSequencia.FieldByName("CONTROLESEQUENCIACONTRATO").AsString <> "S") Then
    Set qSequencia = Nothing
    Exit Function
  End If

  Set qRotina = NewQuery
  Set qContratoIni = NewQuery
  Set qContratoFim = NewQuery

  qContratoIni.Clear

  qContratoIni.Add("SELECT C.CONTRATO  ")
  qContratoIni.Add("  FROM SAM_CONTRATO C")
  qContratoIni.Add(" WHERE C.HANDLE = :HCONTRATOINI ")

  qContratoIni.ParamByName("HCONTRATOINI").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  qContratoIni.Active = True

  qContratoFim.Clear

  qContratoFim.Add("SELECT C.CONTRATO  ")
  qContratoFim.Add("  FROM SAM_CONTRATO C")
  qContratoFim.Add(" WHERE C.HANDLE = :HCONTRATOFIM ")

  qContratoFim.ParamByName("HCONTRATOFIM").AsInteger = CurrentQuery.FieldByName("CONTRATOF").AsInteger
  qContratoFim.Active = True

  qRotina.Clear

  qRotina.Add("SELECT TABTIPOGERACAO ")
  qRotina.Add("  FROM SAM_ROTINACARTAO ")
  qRotina.Add(" WHERE HANDLE = " + CurrentQuery.FieldByName("ROTINACARTAO").AsString)

  qRotina.Active = True

  qSequencia.Clear

  Select Case qRotina.FieldByName("TABTIPOGERACAO").AsInteger
    Case 1
      qSequencia.Add("SELECT REGISTROSEQUENCIALCONTRATO SEQUENCIAL")
      qSequencia.Add("  FROM SAM_CONTRATO")
      qSequencia.Add(" WHERE CONTRATO BETWEEN :CONTRATOINICIAL AND :CONTRATOFINAL")
      qSequencia.Add(" GROUP BY REGISTROSEQUENCIALCONTRATO")

      vMsg = "por Contrato"
    Case 2
      qSequencia.Add("SELECT REGISTROSEQUENCIALAVULSO SEQUENCIAL")
      qSequencia.Add("  FROM SAM_CONTRATO")
      qSequencia.Add(" WHERE CONTRATO BETWEEN :CONTRATOINICIAL AND :CONTRATOFINAL")
      qSequencia.Add(" GROUP BY REGISTROSEQUENCIALAVULSO")

      vMsg = "Avulso"
    Case 3
      qSequencia.Add("SELECT REGISTROSEQUENCIALRENOVACAO SEQUENCIAL")
      qSequencia.Add("  FROM SAM_CONTRATO")
      qSequencia.Add(" WHERE CONTRATO BETWEEN :CONTRATOINICIAL AND :CONTRATOFINAL")
      qSequencia.Add(" GROUP BY REGISTROSEQUENCIALRENOVACAO")

      vMsg = "de Renovação"
    Case 4
      qSequencia.Add("SELECT REGISTROSEQUENCIALSOLICITACAO SEQUENCIAL")
      qSequencia.Add("  FROM SAM_CONTRATO")
      qSequencia.Add(" WHERE CONTRATO BETWEEN :CONTRATOINICIAL AND :CONTRATOFINAL")
      qSequencia.Add(" GROUP BY REGISTROSEQUENCIALSOLICITACAO")

      vMsg = "de Solicitação"
  End Select

  qSequencia.ParamByName("CONTRATOINICIAL").AsInteger = qContratoIni.FieldByName("CONTRATO").AsInteger
  qSequencia.ParamByName("CONTRATOFINAL").AsInteger = qContratoFim.FieldByName("CONTRATO").AsInteger
  qSequencia.Active = True

  'Verificar no intervalo atual, se existem contratos com registros sequenciais diferentes
  vSequenciaAtual = qSequencia.FieldByName("SEQUENCIAL").AsInteger
  vsequencial = qSequencia.FieldByName("SEQUENCIAL").AsInteger

  While (Not qSequencia.EOF)
    If (vsequencial <> qSequencia.FieldByName("SEQUENCIAL").AsInteger) Then
      VerificaSequencial = "Existem Contratos no intervalo com o registro sequencial do cartão (" + vMsg + ") diferentes."

      Set qSequencia = Nothing
      Set qRotina = Nothing
      Set qContratoIni = Nothing
      Set qContratoFim = Nothing

      Exit Function
    End If

    If (qSequencia.FieldByName("SEQUENCIAL").AsInteger = 0) Then
      VerificaSequencial = "Existem Contratos no intervalo sem a definição do registro sequencial do cartão (" + vMsg + ") que será usado na Rotina."

      Set qSequencia = Nothing
      Set qRotina = Nothing
      Set qContratoIni = Nothing
      Set qContratoFim = Nothing

      Exit Function
    End If

    vsequencial = qSequencia.FieldByName("SEQUENCIAL").AsInteger

    qSequencia.Next
  Wend

  'Comparar com os outros intervalos de contratos, se existem contratos com registros sequenciais diferentes
  Dim qAux As Object
  Set qAux = NewQuery

  qAux.Add("SELECT HANDLE, CONTRATO, CONTRATOF ")
  qAux.Add("  FROM SAM_ROTINACARTAO_CONTRATO ")
  qAux.Add(" WHERE ROTINACARTAO = " + CurrentQuery.FieldByName("ROTINACARTAO").AsString)
  qAux.Add("   AND HANDLE <> " + CurrentQuery.FieldByName("HANDLE").AsString)

  qAux.Active = True

  If (Not (qAux.EOF)) Then
    qContratoIni.Active = False

    qContratoIni.ParamByName("HCONTRATOINI").AsInteger = qAux.FieldByName("CONTRATO").AsInteger

    qContratoIni.Active = True

    qContratoFim.Active = False

    qContratoFim.ParamByName("HCONTRATOFIM").AsInteger = qAux.FieldByName("CONTRATOF").AsInteger

    qContratoFim.Active = True

    qSequencia.Active = False

    qSequencia.ParamByName("CONTRATOINICIAL").AsInteger = qContratoIni.FieldByName("CONTRATO").AsInteger
    qSequencia.ParamByName("CONTRATOFINAL").AsInteger = qContratoFim.FieldByName("CONTRATO").AsInteger

    qSequencia.Active = True

    If (vSequenciaAtual <> qSequencia.FieldByName("SEQUENCIAL").AsInteger) Then
      VerificaSequencial = "Existem Contratos em outros intervalos da rotina com o registro sequencial de cartão (" + vMsg + ") diferentes."

      Set qSequencia = Nothing
      Set qRotina = Nothing
      Set qAux = Nothing
      Set qContratoIni = Nothing
      Set qContratoFim = Nothing

      Exit Function
    End If
  End If
End Function

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Select Case CommandID
	Case "BOTAOCANCELAR"
	  BOTAOCANCELAR_OnClick
	Case "BOTAODESBLOQUEAR"
	  BOTAODESBLOQUEAR_OnClick
  End Select
End Sub
