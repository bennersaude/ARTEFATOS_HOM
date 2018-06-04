'HASH: AC668CB74784DB549CF15DBAD87F3BC3
'Macro: SFN_ROTINAARQUIVO
'#Uses "*bsShowMessage"
'#Uses "*CriaTabelaTemporariaSqlServer"
'#Uses "*PrimeiroDiaCompetencia"
' SFN_ROTINAARQUIVO
' alteração: 16/03/00 celso
' -Rotina de cancelamento da rotina arquivo

' Por: Milton
' SMS: 3835
' Sub: TABLE_BeforePost,BOTAOPROCESSAR_OnClick()

'Última alteração: 24/09/2003
'  			  SMS: 18733
'

'A funcao NodeInternalCode é utilizada para determinar se a carga correspondente é da Tarefas de Modelo,
'sendo, mostra o Tab - Modelo para agendamento, não sendo, mostra o Tab - Rotina
'Alteração: 26/12/2005
'      SMS: 52120 - Marcelo Barbosa

'sms 64155 willian
'alterado o tratamento para rotina de cheque


Option Explicit

Public Sub ALTERACONTABLANCDATA
  Dim q As BPesquisa
  Dim Q2 As BPesquisa
  Set q = NewQuery
  Set Q2 = NewQuery
  Dim fin As Object
  Set fin = CreateBennerObject("FINANCEIRO.GERAL")
  Dim CIDADE As Long
  Dim DATA As Date


If NodeInternalCode <> 700 Then

  q.Add("SELECT MUNICIPIOPADRAO FROM SFN_PARAMETROSFIN")
  q.Active = True
  CIDADE = q.FieldByName("MUNICIPIOPADRAO").AsInteger

  q.Clear
  q.Add("SELECT DISTINCT DATA FROM SFN_CONTAB_LANC WHERE DATA>=:DATAI AND DATA<=:DATAF AND ROTINAARQUIVO IS NULL ")
  q.ParamByName("DATAI").Value = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
  q.ParamByName("DATAF").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
  q.Active = True

  If Not InTransaction Then StartTransaction

  Q2.Clear
  Q2.Add("UPDATE SFN_CONTAB_LANC SET DATA=:DATANOVA WHERE DATA=:DATAVELHA AND  ROTINAARQUIVO IS NULL")

  While Not q.EOF
    DATA = fin.PROXDIAUTIL(CurrentSystem, q.FieldByName("DATA").AsDateTime, CIDADE)
    If DATA <>q.FieldByName("DATA").AsDateTime Then
      Q2.ParamByName("DATAVELHA").Value = q.FieldByName("DATA").AsDateTime
      Q2.ParamByName("DATANOVA").Value = DATA
      Q2.ExecSQL
    End If

    q.Next
  Wend

  If InTransaction Then Commit

End If

  Set fin = Nothing
  Set q = Nothing
  Set Q2 = Nothing

End Sub

Public Sub BOTAOAGENDAR_OnClick()
  Dim qr As BPesquisa
  Dim qr1 As BPesquisa
  Dim vSituacao As String
  Dim vTabela As String
  Dim vLegendaAgendamento As String
  Dim VLegendaAberta As String
  Dim VLegendaProcessada As String
  Set qr = NewQuery
  Set qr1 = NewQuery
  vTabela = "SFN_ROTINAARQUIVO"
  vLegendaAgendamento = "3"
  VLegendaAberta = "1"
  VLegendaProcessada = "5"
  qr.Clear
  qr.Add("SELECT SITUACAO FROM " + vTabela + " WHERE HANDLE = :pHANDLE")
  qr.ParamByName("pHandle").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qr.Active = True
  vSituacao = qr.FieldByName("SITUACAO").AsString
  If vSituacao <> vLegendaAgendamento Then
    If vSituacao = VLegendaAberta Then
      If bsShowMessage("Confirme o agendamento da rotina", "Q") = vbYes Then '(6=yes, 7=não)
        qr1.Clear
        qr1.Add("UPDATE " + vTabela + " SET SITUACAO = :pSituacao WHERE HANDLE = :pHANDLE")
        qr1.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
        qr1.ParamByName("pSituacao").AsString = vLegendaAgendamento
        qr1.ExecSQL
      End If
    Else
      bsShowMessage("Rotina já foi processada.", "I")
    End If
  Else
    If bsShowMessage("Rotina já está agendada. Para retirar o agendamento pressione 'SIM'", "Q") = vbYes Then
      qr1.Clear
      qr1.Add("UPDATE " + vTabela + " SET SITUACAO = :pSituacao WHERE HANDLE = :pHANDLE")
      qr1.ParamByName("pHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      If (CurrentQuery.FieldByName("PROCESSADODATA").IsNull) Then
        qr1.ParamByName("pSituacao").AsString = VLegendaAberta
      Else
        qr1.ParamByName("pSituacao").AsString = VLegendaProcessada
      End If
      qr1.ExecSQL
    End If
  End If
  Set qr = Nothing
  Set qr1 = Nothing
  SelectNode(CurrentQuery.FieldByName("HANDLE").AsInteger, True, False)
End Sub

Public Sub BOTAOCONFIRMAR_OnClick()
  Dim vsMensagem As String
  Dim viRetorno As Long
  Dim vcContainer As CSDContainer

  'FAZ A BAIXA DOS DOCUMENTOS SE FOR CHEQUE =A ROTINA DE RETORNO
  If CurrentQuery.FieldByName("CONFIRMADODATA").IsNull Then
    If CurrentQuery.FieldByName("PROCESSADODATA").IsNull Then
      bsShowMessage("A rotina ainda não foi processada.", "I")
    Else
      If (CurrentQuery.FieldByName("TABTIPO").AsInteger <>2) _
        Or(CurrentQuery.FieldByName("TABTIPO").AsInteger <> 7) Then

         If(Not CurrentQuery.FieldByName("TESOURARIA").IsNull) _
           Or(CurrentQuery.FieldByName("TABTIPO").AsInteger = 4) _
           Or(CurrentQuery.FieldByName("TABTIPO").AsInteger = 5) Then

             If((CurrentQuery.FieldByName("TABTIPO").AsInteger = 4) _
               And(CurrentQuery.FieldByName("PASTACORPORATIVO").IsNull) _
               And(CurrentQuery.FieldByName("TABARQUIVO").AsInteger = 2)) Then
                 bsShowMessage("Campo Pasta deve ser preenchido.", "I")
             Else
              Dim Interface As Object

			  If VisibleMode Then
      			If Not VerificaArquivo Then 'salva o arquivo de retorno no campo do tipo ARQUIVO .
      				Exit Sub
      			End If
  			    Set Interface = CreateBennerObject("BSINTERFACE0030.RotinaArquivo")
  				Interface.Confirma(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  		        CurrentQuery.Active = False
                CurrentQuery.Active = True
			  Else
  				Set Interface = CreateBennerObject("BSServerExec.ProcessosServidor")
  				viRetorno = Interface.ExecucaoImediata(CurrentSystem, _
                                                       "RotArq", _
                                         			   "RotinaArquivo_ConfirmaRotina", _
                                         			   "Rotina Arquivo - Confirmação", _
                                                       CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                                       "SFN_ROTINAARQUIVO", _
                                                       "SITUACAO", _
                                                       "", _
                                                       "", _
                                                       "P", _
                                                       True, _
                                                       vsMensagem, _
                                                       Null)

  				If viRetorno = 0 Then
                  bsShowMessage("Processo enviado para execução no servidor!", "I")
                Else
                  bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagem, "I")
                End If
              End If
              Set Interface = Nothing
          End If
        Else
          bsShowMessage("A tesouraria deve ser informada.", "I")
        End If
      Else
        bsShowMessage("Rotina de confirmação somente para cheques e arquivo SIC.", "I")
      End If
    End If
  Else
    bsShowMessage("A rotina já foi confirmada.", "I")
  End If

End Sub

Public Sub BOTAOIMPRIMERECIBO_OnClick()

  Dim SelecionaHandleRelatorio As BPesquisa
  Set SelecionaHandleRelatorio = NewQuery

  If bsShowMessage("Deseja visualizar os recibos?", "Q") = vbOK Then

    SelecionaHandleRelatorio.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO =:PCODIGO")
    SelecionaHandleRelatorio.ParamByName("PCODIGO").AsString = "SFN-RA009"
    SelecionaHandleRelatorio.Active = False
    SelecionaHandleRelatorio.Active = True

    ReportPreview(SelecionaHandleRelatorio.FieldByName("HANDLE").AsInteger, "EXISTS (SELECT HANDLE FROM SFN_ROTINAARQUIVO_DOC WHERE A.HANDLE = DOCUMENTO AND ROTINAARQUIVO =" + CurrentQuery.FieldByName("HANDLE").AsString + ")", False, True)
  End If
  Set SelecionaHandleRelatorio = Nothing

End Sub


Public Sub BOTAOINTEGRACAOSIAFI_OnClick()

  If (CurrentQuery.State = 2) Or (CurrentQuery.State = 3) Then
  	bsShowMessage("Ação não permitida com a rotina em edição ou inclusão!", "I")
  	Exit Sub
  End If

  If (CurrentQuery.FieldByName("TABTIPO").AsInteger = 1) Then
    If (CurrentQuery.FieldByName("SITUACAO").AsString <> "5") Then
      bsShowMessage("Ação permitida apenas para rotinas de remessa já processadas!", "I")
  	  Exit Sub
    End If

    If (CurrentQuery.FieldByName("SEQUENCIALSIAFI").IsNull) Then
      Dim contador As Long
      contador = 0
	  NewCounter2("SIAFI", 0, 1, contador)

	  Dim qSql As BPesquisa
      Set qSql = NewQuery
      qSql.Add("UPDATE SFN_ROTINAARQUIVO             ")
      qSql.Add("   SET SEQUENCIALSIAFI = :SEQUENCIAL ")
      qSql.Add(" WHERE HANDLE = :HANDLE              ")

      qSql.ParamByName("SEQUENCIAL").AsInteger = contador
      qSql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
      qSql.ExecSQL

      qSql.Active = False
      Set qSql = Nothing
	End If


    Dim IntegracaoSiafi As CSBusinessComponent
    Set IntegracaoSiafi = BusinessComponent.CreateInstance("Benner.Saude.Financeiro.Business.SfnRotinaArquivoBLL, Benner.Saude.Financeiro.Business")
    IntegracaoSiafi.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
    IntegracaoSiafi.Execute("GerarArquivoXmlSiafi")
    Set IntegracaoSiafi = Nothing
    bsShowMessage("Processo concluído, verifique os arquivos na carga 'Arquivos SIAFI'!", "I")
  Else
    If (CurrentQuery.FieldByName("TABTIPO").AsInteger = 2) Then
      bsShowMessage("O Retorno deve ser executado atravez do botão 'Processar'!", "I")
    Else
      bsShowMessage("Ação permitida apenas para a rotina de Remessa!", "I")
  	End If
  End If
End Sub

Public Sub BOTAOMODELORELATORIO_OnClick()
  Dim Interface As Object
  Set Interface = CreateBennerObject("rotarq.rotinas")
  Interface.ModeloRelatorio(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set Interface = Nothing
End Sub

Public Sub BOTAOPROCESSAR_OnClick()
  Dim Interface As Object
  Dim Retorno As Boolean
  Dim QCampos As BPesquisa
  Dim vHnd As Long
  Dim vbbjm As Boolean
  'SMS 90427 - Marcelo Barbosa - 18/03/2008
  Dim QCheque As BPesquisa
  Dim vsMensagem As String
  Dim viRetorno As Long
  Dim vcContainer As CSDContainer
  Dim qParam As Object
  Dim PERIODOFATCONINICIAL As Date
  Dim PERIODOFATCONFINAL   As Date

  If (CurrentQuery.State = 2) Or (CurrentQuery.State = 3) Then
  	bsShowMessage("A rotina não pode ser processada em edição ou inclusão!", "I")
  	Exit Sub
  End If

  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 4 Then

    Set qParam = NewQuery

    qParam.Clear
    qParam.Add("SELECT PERIODOFATCONINICIAL, PERIODOFATCONFINAL,  CONTABILIZA FROM SFN_PARAMETROSFIN")
    qParam.Active = True

    If qParam.FieldByName("CONTABILIZA").AsString = "S" Then
      PERIODOFATCONINICIAL = PrimeiroDiaCompetencia(qParam.FieldByName("PERIODOFATCONINICIAL").AsDateTime)

	  If (qParam.FieldByName("PERIODOFATCONFINAL").IsNull) Then
        If Not ((CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >= PERIODOFATCONINICIAL) And (PERIODOFATCONINICIAL <= CurrentQuery.FieldByName("DATAFINAL").AsDateTime)) Then
           bsShowMessage("Não é permitido Processar uma rotina cuja data inicial/final esteja fora do período contábil.", "E")
           Set qParam = Nothing
           Exit Sub
        End If
	  Else
	    PERIODOFATCONFINAL = PrimeiroDiaCompetencia(qParam.FieldByName("PERIODOFATCONFINAL").AsDateTime)
        If Not ((CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >= PERIODOFATCONINICIAL) And (CurrentQuery.FieldByName("DATAFINAL").AsDateTime <= PERIODOFATCONFINAL)) Then
           bsShowMessage("Não é permitido Processar uma rotina cuja data inicial/final esteja fora do período contábil.", "E")
           Set qParam = Nothing
           Exit Sub
        End If
      End If
    End If

    Set qParam = Nothing
  End If

  Set QCampos = NewQuery
  Retorno = True 'Devera ser verdadeira,devido ao processamento de outras rotinas que não sejam Interface Cliente(Tab 6).

  ' SMS 40765 lOPES
  QCampos.Active = False
  QCampos.Clear
  QCampos.Add("SELECT HANDLE ")
  QCampos.Add("  FROM SFN_MODELO_ESTRUTURA ")
  QCampos.Add(" WHERE MODELO = :MOD")
  QCampos.ParamByName("MOD").AsInteger = CurrentQuery.FieldByName("MODELO").AsInteger
  QCampos.Active = True
  vHnd = QCampos.FieldByName("HANDLE").AsInteger



  QCampos.Active = False
  QCampos.Clear
  QCampos.Add("Select COUNT(HANDLE) QT ")
  QCampos.Add("  FROM SFN_MODELO_ESTRUTURA_CAMPO ")
  QCampos.Add(" WHERE MODELOESTRUTURA = :MOD")
  QCampos.Add("   And CAMPO In (SELECT HANDLE ")
  QCampos.Add("                   FROM SIS_CONTABCAMPOS ")
  QCampos.Add("                  WHERE NOME = 'BAIXAJUROSMULTA' ")
  QCampos.Add("                     OR NOME = 'BAIXAMULTA') ")
  QCampos.ParamByName("MOD").AsInteger = vHnd
  QCampos.Active = True
  vbbjm = False
  If QCampos.FieldByName("QT").AsInteger = 2 Then
    vbbjm = True
  End If

  QCampos.Active = False
  QCampos.Clear
  QCampos.Add("Select COUNT(HANDLE) QT ")
  QCampos.Add("  FROM SFN_MODELO_ESTRUTURA_CAMPO ")
  QCampos.Add(" WHERE MODELOESTRUTURA = :MOD")
  QCampos.Add("   And CAMPO In (SELECT HANDLE ")
  QCampos.Add("                   FROM SIS_CONTABCAMPOS ")
  QCampos.Add("                  WHERE NOME = 'BAIXAJUROSMULTA' ")
  QCampos.Add("                     OR NOME = 'BAIXAJURO') ")
  QCampos.ParamByName("MOD").AsInteger = vHnd
  QCampos.Active = True
  If QCampos.FieldByName("QT").AsInteger = 2 Then
    vbbjm = True
  End If

  Set QCampos = Nothing

  If vbbjm Then
    bsShowMessage("O modelo não pode possuir o campo 'Baixa juros e multa' juntamente com o campo 'Baixa juro' ou com o campo 'Baixa multa'", "I")
    Exit Sub
  End If
  'SMS 90427 - Marcelo Barbosa - 18/03/2008
  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 3 Then

    Set QCheque = NewQuery
    QCheque.Active = False
    QCheque.Add("SELECT NUMEROCHEQUEDISPONIVEL FROM SFN_TESOURARIA WHERE HANDLE=:TESOURARIA")
    QCheque.ParamByName("TESOURARIA").AsInteger = CurrentQuery.FieldByName("TESOURARIA").AsInteger
    QCheque.Active = True
    If QCheque.FieldByName("NUMEROCHEQUEDISPONIVEL").AsInteger <> CurrentQuery.FieldByName("NUMEROCHEQUE").AsInteger Then
      If bsShowMessage("Número do cheque na rotina difere da tesouraria. Continuar?", "Q") <> vbYes Then
         bsShowMessage("'Cancelado pelo usuário! (Problemas com o número do cheque)" , "I")
         Exit Sub
      End If
    End If
    Set QCheque = Nothing
  End If
  'Fim - SMS 90427

  If CurrentQuery.FieldByName("TABTIPO").AsInteger < 1001 Then 'Quando for 1001 (Interface Cliente) o processo será executado na camada específica
    If VisibleMode Then
      If Not VerificaArquivo Then 'salva o arquivo de retorno no campo do tipo ARQUIVO .
        Exit Sub
      End If
      Set Interface = CreateBennerObject("BSINTERFACE0030.RotinaArquivo")
      Interface.Processa(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
      CurrentQuery.Active = False
      CurrentQuery.Active = True
    Else
	  Set vcContainer = NewContainer

	  vcContainer.AddFields("HANDLE:INTEGER")
	  vcContainer.Insert
	  vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

      Set Interface = CreateBennerObject("BSServerExec.ProcessosServidor")
      viRetorno = Interface.ExecucaoImediata(CurrentSystem, _
                                      "RotArq", _
                                      "RotinaArquivo_ProcessaRotina", _
                                      "Rotina Arquivo - Processamento", _
                                      CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                      "SFN_ROTINAARQUIVO", _
                                      "SITUACAO", _
                                      "", _
                                      "", _
                                      "P", _
                                      False, _
                                      vsMensagem, _
                                      vcContainer)

      If viRetorno = 0 Then
        bsShowMessage("Processo enviado para execução no servidor!", "I")
      Else
        bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagem, "I")
      End If

      Set Interface = Nothing
	  Set vcContainer = Nothing
    End If
  End If
End Sub


Public Sub BOTAOCANCELAR_OnClick()
  Dim vcContainer As CSDContainer

  If (CurrentQuery.State = 2) Or (CurrentQuery.State = 3) Then
  	bsShowMessage("A rotina não pode ser cancelada em edição ou inclusão!", "I")
  	Exit Sub
  End If

  If bsShowMessage("Confirma o cancelamento da rotina ?", "Q") = vbYes Then
    Dim Interface As Object
    Dim viRetorno As Integer
    Dim vsMensagem As String

    If VisibleMode Then
      Set Interface = CreateBennerObject("BSINTERFACE0030.RotinaArquivo")
      Interface.Cancela(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
      CurrentQuery.Active = False
      CurrentQuery.Active = True
    Else
      Set vcContainer = NewContainer

	  vcContainer.AddFields("HANDLE:INTEGER")
	  vcContainer.Insert
	  vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

	  Set Interface = CreateBennerObject("BSServerExec.ProcessosServidor")
      viRetorno = Interface.ExecucaoImediata(CurrentSystem, _
                                   "RotArq", _
                                   "RotinaArquivo_CancelaRotina", _
                                   "Rotina Arquivo - Cancelamento", _
                                   CurrentQuery.FieldByName("HANDLE").AsInteger, _
                                   "SFN_ROTINAARQUIVO", _
                                   "SITUACAO", _
                                   "", _
                                   "", _
                                   "C", _
                                   True, _
                                   vsMensagem, _
                                   vcContainer)
      If viRetorno = 0 Then
        bsShowMessage("Processo enviado para execução no servidor!", "I")
      Else
        bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagem, "I")
      End If
    End If
    Set interface = Nothing
	Set vcContainer = Nothing
  End If
End Sub

Public Sub BOTAOSIAFIORDEMBANCARIA_AfterOnClick()
RefreshNodesWithTable "SFN_ROTINAARQUIVO"
End Sub

Public Sub BOTAOSIAFIORDEMBANCARIA_OnClick()
  Dim interface As Object
  Dim viRetorno As Integer
  Dim vsMensagem As String
  Dim vvContainer As CSDContainer

   Set vvContainer = NewContainer

  UserVar("HANDLE_ROTINAARQUIVO") = CurrentQuery.FieldByName("HANDLE").AsString

  Set interface = CreateBennerObject("BSInterface0002.GerarFormularioVirtual")
  viRetorno = interface.Exec(CurrentSystem, _
                             1, _
                             "TV_SIAFIORDEMBANCARIA", _
                             "Alterar Ordem Bancária", _
                             0, _
                             450, _
                             400, _
                             False, _
                             vsMensagem, _
                             vvContainer)

  Set interface =Nothing
  Set vvContainer = Nothing

  If vsMensagem <> "" Then
      bsShowmessage(vsMensagem, "E")
  End If
End Sub

Public Sub INTEGRACAOCOMCOMPRASDE_OnChange()
  If AUTORIZACAO.Visible Then
  	CurrentQuery.FieldByName("AUTORIZACAO").Clear
  End If
  CurrentQuery.FieldByName("FORNECIMENTO").Clear
  CurrentQuery.FieldByName("MODELO").Clear
End Sub

Public Sub MODELO_OnChange()
  CurrentQuery.FieldByName("TIPODOCUMENTO").Clear
End Sub

Public Sub MODELO_OnPopup(ShowPopup As Boolean)

  If TABTIPO.PageIndex =3 Then
     MODELO.LocalWhere ="SFN_MODELO.TABTIPO=3"

  ElseIf TABTIPO.PageIndex =4 Then
     MODELO.LocalWhere ="SFN_MODELO.TABTIPO=9"

  ElseIf TABTIPO.PageIndex = 5 Then
     MODELO.LocalWhere ="SFN_MODELO.TABTIPO = 10"

  ElseIf TABTIPO.PageIndex = 6 Then
     MODELO.LocalWhere ="SFN_MODELO.TABTIPO = 12 AND SFN_MODELO.INTEGRACAOCOMCOMPRASDE = '" + CurrentQuery.FieldByName("INTEGRACAOCOMCOMPRASDE").AsString + "'" 'IntegraÃ§Ã£o com o compras

  Else
     MODELO.LocalWhere ="SFN_MODELO.TABTIPO<>3"
  End If
End Sub

Public Sub TABLE_AfterInsert()
  TABTIPO_OnChange
End Sub

Public Sub TABLE_AfterPost()
  If CurrentQuery.FieldByName("TABTIPO").AsInteger <> 2 And CurrentQuery.FieldByName("ARQUIVOSIAFI").AsInteger <> 1 Then
	VerificaArquivo
  End If
End Sub

Public Sub TABLE_AfterScroll()

	SessionVar("HNDLROTARQ") = CStr(CurrentQuery.FieldByName("HANDLE").AsInteger)

	BOTAOAGENDAR.Visible = True
	BOTAOCANCELAR.Visible = True
	BOTAOCONFIRMAR.Visible = True
	BOTAOIMPRIMERECIBO.Visible = True
	BOTAOPROCESSAR.Visible = True
	BOTAORELATORIO.Visible = True
	If CurrentQuery.FieldByName("TABTIPO").Value = 5 Then 'Tesouraria
		BOTAORELATORIO.Visible = False
	End If

	If(CurrentQuery.FieldByName("TABTIPO").AsInteger =7)Or(CurrentQuery.FieldByName("TABTIPO").AsInteger =2)Or(CurrentQuery.FieldByName("TABTIPO").AsInteger = 1001)Then
		BOTAOCONFIRMAR.Enabled =False
	Else
		BOTAOCONFIRMAR.Enabled =True
	End If
	If CurrentQuery.FieldByName("CANCELADODATA").IsNull Then
		BOTAOCANCELAR.Enabled =True
	Else
		BOTAOCANCELAR.Enabled =False
	End If
	BOTAOCANCELAR.Enabled =True
	If BOTAOCONFIRMAR.Visible = True Then
		If(CurrentQuery.FieldByName("TABTIPO").AsInteger =4)Or(CurrentQuery.FieldByName("TABTIPO").AsInteger =5)Then
			BOTAOCONFIRMAR.Caption ="Gerar Arquivo"
		Else
			BOTAOCONFIRMAR.Caption ="Confirmar"
		End If
	End If
	If CurrentQuery.FieldByName("TABTIPO").AsInteger =1 Or CurrentQuery.FieldByName("TABTIPO").AsInteger =3 Then
		BOTAOIMPRIMERECIBO.Visible =True
	Else
		BOTAOIMPRIMERECIBO.Visible =False
	End If

	If CurrentQuery.FieldByName("TABREMESSARETORNOOPME").AsInteger = 2 Then
		AUTORIZACAO.ReadOnly = True
		FORNECIMENTO.ReadOnly = True
  	Else
		AUTORIZACAO.ReadOnly = False
    	FORNECIMENTO.ReadOnly = False
  	End If

	If CurrentQuery.FieldByName("TABTIPO").AsInteger = 2 And CurrentQuery.FieldByName("ARQUIVOSIAFI").AsInteger = 1 And CurrentQuery.FieldByName("SITUACAO").AsString = "5" Then
		BOTAOPROCESSAR.Enabled = False
	Else
		BOTAOPROCESSAR.Enabled = True
	End If


    OcultaBotaoIntegracaoSiafi
End Sub



Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If WebMode Then
		If(Not CurrentQuery.FieldByName("PROCESSADODATA").IsNull) And CurrentQuery.FieldByName("TABTIPO").AsInteger <> 4 Then
			bsShowMessage("A rotina está processada!", "E")
	  		CanContinue =False
    	End If
	Else
	  	If NodeInternalCode <> 700 Then
    		If(Not CurrentQuery.FieldByName("PROCESSADODATA").IsNull) And CurrentQuery.FieldByName("TABTIPO").AsInteger <> 4 Then
	  			CanContinue =False
	   			bsShowMessage("A rotina está processada!", "E")
    		End If
  		End If
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vDllSfnRotinaArquivo As Object
  Set vDllSfnRotinaArquivo = CreateBennerObject("RotArq.SfnRotinaArquivo")

  CanContinue = vDllSfnRotinaArquivo.BeforePost(CurrentSystem)
  Set vDllSfnRotinaArquivo = Nothing

  If CurrentQuery.FieldByName("TABTIPO").AsInteger = 2 And CurrentQuery.FieldByName("ARQUIVOSIAFI").AsInteger = 1 Then
	If CurrentQuery.FieldByName("NUMERODOCUMENTOHABIL").IsNull And CurrentQuery.FieldByName("VENCIMENTOINICIAL").IsNull And CurrentQuery.FieldByName("VENCIMENTOFINAL").IsNull Then
	  bsShowMessage("Dever ser informado Número do documento hábil ou data de vencimento!", "E")
	  CanContinue = False
    Else
      If Not CurrentQuery.FieldByName("VENCIMENTOINICIAL").IsNull And CurrentQuery.FieldByName("VENCIMENTOFINAL").IsNull Then
	    bsShowMessage("Falta informar data final do vencimento!", "E")
	    CanContinue = False
	  ElseIf Not CurrentQuery.FieldByName("VENCIMENTOFINAL").IsNull And CurrentQuery.FieldByName("VENCIMENTOINICIAL").IsNull  Then
	    bsShowMessage("Falta informar data inicial do vencimento!", "E")
	    CanContinue = False
	  End If
	End If
	If CurrentQuery.FieldByName("ARQUIVORETORNO").AsString <> "" Then
	  bsShowMessage("Para Arquivo SIAFI não incluir arquivo retorno", "E")
	  CanContinue = False
	End If
  End If


  If (CanContinue = False) Then
    Exit Sub
  End If

End Sub

Public Sub TABLE_NewRecord()
  'SMS 90427 - Marcelo Barbosa - 23/04/2008
  CurrentQuery.FieldByName("TABTIPO").Value = 1

  BOTAOAGENDAR.Visible = True
  BOTAOCANCELAR.Visible = True
  BOTAOCONFIRMAR.Visible = True
  BOTAOIMPRIMERECIBO.Visible = True
  BOTAOPROCESSAR.Visible = True
  BOTAORELATORIO.Visible = True

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOAGENDAR"
			BOTAOAGENDAR_OnClick
		Case "BOTAOCONFIRMAR"
			BOTAOCONFIRMAR_OnClick
		Case "BOTAOIMPRIMERECIBO"
			BOTAOIMPRIMERECIBO_OnClick
		Case "BOTAOMODELORELATORIO"
			BOTAOMODELORELATORIO_OnClick
		Case "BOTAOPROCESSAR"
			BOTAOPROCESSAR_OnClick
		Case "K9_BOTAOPROCESSAR"
            BOTAOPROCESSAR_OnClick
		Case "BOTAOCANCELAR"
			BOTAOCANCELAR_OnClick
		Case "K9_BOTAOCANCELAR"
            BOTAOCANCELAR_OnClick
	End Select
End Sub

Public Sub TABREMESSARETORNOOPME_OnChange()

  If TABREMESSARETORNOOPME.PageIndex = 1 Then
	AUTORIZACAO.ReadOnly = True
	FORNECIMENTO.ReadOnly = True
  Else
	AUTORIZACAO.ReadOnly = False
    FORNECIMENTO.ReadOnly = False
  End If

End Sub

Public Sub TABTIPO_OnChange()
  If CurrentQuery.State <> 1 Then
    If TABTIPO.PageIndex = 5 Then
      CurrentQuery.FieldByName("TABTIPO").Value = 7
    End If
  End If

  If (CurrentQuery.FieldByName("PROCESSADODATA").IsNull) Then
    If TABTIPO.PageIndex = 0 Then
      If CurrentQuery.FieldByName("VENCIMENTOINICIAL").IsNull Then
        CurrentQuery.FieldByName("VENCIMENTOINICIAL").Value = CurrentQuery.FieldByName("DATAARQUIVO").AsDateTime
      End If

      If CurrentQuery.FieldByName("VENCIMENTOFINAL").IsNull Then
        CurrentQuery.FieldByName("VENCIMENTOFINAL").Value = CurrentQuery.FieldByName("DATAARQUIVO").AsDateTime
      End If
    Else
      CurrentQuery.FieldByName("VENCIMENTOINICIAL").Clear
      CurrentQuery.FieldByName("VENCIMENTOFINAL").Clear
    End If
  End If

  OcultaBotaoIntegracaoSiafi(TABTIPO.PageIndex)

End Sub

Public Sub BOTAORELATORIO_OnClick()
  Dim sql As BPesquisa
  Set sql = NewQuery

  sql.Add("SELECT HANDLE FROM R_RELATORIOS WHERE CODIGO = 'CRE003'")
  sql.Active = True

  ReportPreview(sql.FieldByName("HANDLE").Value, "", True, False)

  Set sql = Nothing

End Sub

Public Sub TESOURARIA_OnChange()
'SMS 52120 - Marcelo Barbosa - 26/12/2005
  Dim sql As BPesquisa
  Set sql = NewQuery


  sql.Add("SELECT NUMEROCHEQUEDISPONIVEL FROM SFN_TESOURARIA WHERE HANDLE = :HANDLE")
  sql.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("TESOURARIA").AsInteger
  sql.Active = True

  If (sql.FieldByName("NUMEROCHEQUEDISPONIVEL").AsInteger) > 0 Then
	CurrentQuery.FieldByName("NUMEROCHEQUE").AsInteger = sql.FieldByName("NUMEROCHEQUEDISPONIVEL").AsInteger
  Else
	CurrentQuery.FieldByName("NUMEROCHEQUE").Clear
  End If

  sql.Active=False
  Set sql = Nothing
End Sub

Public Sub TIPODOCUMENTO_OnChange()
  If CurrentQuery.FieldByName("TABTIPO").AsInteger <> 7 Then 'Rotina diferente de corporativo
    TIPODOCUMENTO.LocalWhere = "HANDLE IN (SELECT TD.TIPODOCUMENTO FROM SFN_TIPODOCUMENTO_MODELO TD WHERE TD.MODELODOCUMENTO = " + Str(CurrentQuery.FieldByName("MODELO").AsInteger) + ")"
  Else
    TIPODOCUMENTO.LocalWhere = ""
  End If

End Sub

Public Sub TIPODOCUMENTO_OnPopup(ShowPopup As Boolean)
  If CurrentQuery.FieldByName("TABTIPO").AsInteger <> 7 Then 'Rotina diferente de corporativo
    TIPODOCUMENTO.LocalWhere = "HANDLE IN (SELECT TD.TIPODOCUMENTO FROM SFN_TIPODOCUMENTO_MODELO TD WHERE TD.MODELODOCUMENTO = " + Str(CurrentQuery.FieldByName("MODELO").AsInteger) + ")"
  Else
    TIPODOCUMENTO.LocalWhere = ""
  End If
End Sub

Public Sub AUTORIZACAO_OnChange()
  If CurrentQuery.FieldByName("INTEGRACAOCOMCOMPRASDE").AsString = "M" Then
    AUTORIZACAO.LocalWhere = "HANDLE In (Select AUTORIZACAO " + _
             " FROM SAM_AUTORIZ_FORNECIMENTO F " + _
            " WHERE F.SITUACAO = '4' " + _
             " And F.TIPOFORNECIMENTOPAI In (Select HANDLE FROM SAM_TIPOFORNECIMENTO TF" + _
                            				" WHERE TF.EXIGETIPOAQUISICAO = 'S')) "
  Else
    AUTORIZACAO.LocalWhere = "HANDLE In (Select AUTORIZACAO " + _
             "FROM SAM_AUTORIZ_FORNECIMENTO F" + _
            "WHERE F.SITUACAOANALISETECNICA = '4'" + _
             " And F.Handle In (Select FORNECIMENTO FROM SAM_AUTORIZ_FORNEC_EVENTO FE " + _
                               " WHERE FE.TIPOFORNECIMENTO In (Select HANDLE FROM SAM_TIPOFORNECIMENTO TF " + _
                                                             " WHERE TF.DEFINIDOPELODEPTOCOMPRAS = 'S')) " + _
                                ")"
  End If
End Sub

Public Sub AUTORIZACAO_OnPopup(ShowPopup As Boolean)
  If CurrentQuery.FieldByName("INTEGRACAOCOMCOMPRASDE").AsString = "M" Then
    AUTORIZACAO.LocalWhere = "HANDLE In (Select AUTORIZACAO " + _
             "FROM SAM_AUTORIZ_FORNECIMENTO F " + _
            "WHERE F.SITUACAO = '4' " + _
             " And F.TIPOFORNECIMENTOPAI In (Select HANDLE FROM SAM_TIPOFORNECIMENTO TF" + _
                 "                            WHERE TF.EXIGETIPOAQUISICAO = 'S')) "
  Else
    AUTORIZACAO.LocalWhere = "HANDLE In (Select AUTORIZACAO " + _
             " FROM SAM_AUTORIZ_FORNECIMENTO F " + _
            " WHERE F.SITUACAOANALISETECNICA = '4'" + _
             " And F.Handle In (Select FORNECIMENTO FROM SAM_AUTORIZ_FORNEC_EVENTO FE " + _
                               " WHERE FE.TIPOFORNECIMENTO In (Select HANDLE FROM SAM_TIPOFORNECIMENTO TF " + _
                                 "                              WHERE TF.DEFINIDOPELODEPTOCOMPRAS = 'S')) " + _
                                ")"
  End If
End Sub

Public Sub FORNECIMENTO_OnChange()
  If CurrentQuery.FieldByName("INTEGRACAOCOMCOMPRASDE").AsString = "M" Then
    FORNECIMENTO.LocalWhere = "SITUACAO = '4'" + _
                          "And TIPOFORNECIMENTOPAI In (Select TF.Handle FROM SAM_TIPOFORNECIMENTO TF " + _
                                                      " WHERE TF.EXIGETIPOAQUISICAO = 'S')"
  Else
    FORNECIMENTO.LocalWhere = "SITUACAOANALISETECNICA = '4' " + _
    "And HANDLE In (Select FE.FORNECIMENTO FROM SAM_AUTORIZ_FORNEC_EVENTO FE" + _
                    "WHERE FE.TIPOFORNECIMENTO In (Select HANDLE FROM SAM_TIPOFORNECIMENTO TF " + _
                                                   "WHERE TF.DEFINIDOPELODEPTOCOMPRAS = 'S')) "
  End If
End Sub

Public Sub FORNECIMENTO_OnPopup(ShowPopup As Boolean)
  If CurrentQuery.FieldByName("INTEGRACAOCOMCOMPRASDE").AsString = "M" Then
    FORNECIMENTO.LocalWhere = "SITUACAO = '4' " + _
                         " And TIPOFORNECIMENTOPAI In (Select TF.Handle FROM SAM_TIPOFORNECIMENTO TF " + _
                                                      " WHERE TF.EXIGETIPOAQUISICAO = 'S')"
  Else
    FORNECIMENTO.LocalWhere = "SITUACAOANALISETECNICA = '4' " + _
                         "And HANDLE In (Select FE.FORNECIMENTO FROM SAM_AUTORIZ_FORNEC_EVENTO FE                  " + _
                                        " WHERE FE.TIPOFORNECIMENTO In (Select HANDLE FROM SAM_TIPOFORNECIMENTO TF " + _
                                                                       " WHERE TF.DEFINIDOPELODEPTOCOMPRAS = 'S')) "
  End If
End Sub

Function VerificaArquivo As Boolean
	VerificaArquivo = True
	If CurrentQuery.FieldByName("TABTIPO").AsInteger  = 2 Then ' RETORNO
		If Trim(CurrentQuery.FieldByName("ARQUIVO").AsString) <> "" Then
			If Dir(CurrentQuery.FieldByName("ARQUIVO").AsString,vbArchive) <> "" Then
				SetFieldDocument("SFN_ROTINAARQUIVO","ARQUIVORETORNO",CurrentQuery.FieldByName("HANDLE").AsInteger,CurrentQuery.FieldByName("ARQUIVO").AsString,True)
			Else
				bsShowMessage("Não foi possível localizar o arquivo de retorno!", "I")
				VerificaArquivo = False
			End If
		End If
	End If
End Function

Function UsaDotacaoOrcamentaria As Boolean
  Dim qSql As BPesquisa
  Set qSql = NewQuery
  qSql.Add(" SELECT COUNT(1) EXISTE    ")
  qSql.Add("   FROM SFN_PARAMETROSFIN  ")
  qSql.Add("  WHERE CONTROLADOTORC = 2 ")

  qSql.Active = True

  If qSql.FieldByName("EXISTE").AsInteger = 1 Then
    UsaDotacaoOrcamentaria = True
  Else
    UsaDotacaoOrcamentaria = False
  End If

  qSql.Active = False
  Set qSql = Nothing
End Function

Public Sub OcultaBotaoIntegracaoSiafi(Optional pTABTIPO As Integer = -1)
  Dim vTabRemessa As Boolean

  If pTABTIPO >= 0 Then
    vTabRemessa = pTABTIPO = 0
  Else
    vTabRemessa = CurrentQuery.FieldByName("TABTIPO").AsInteger = 1
  End If

  BOTAOINTEGRACAOSIAFI.Visible = (vTabRemessa) And (UsaDotacaoOrcamentaria)
End Sub
