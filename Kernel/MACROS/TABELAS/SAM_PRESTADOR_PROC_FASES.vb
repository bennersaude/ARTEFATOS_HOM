'HASH: D2088FE76F49F01DF7E723AFA2208F44
'#Uses "*bsShowMessage"

'Mauricio Ibelli - 04/01/2002 - sms3165 - Se filial padrao do prestador for nulo não checar responsavel

Option Explicit

Dim Mensagem As String
Dim vDataInicial As Date
Dim podeContinuar As Boolean
Dim executadoPelaWeb As Boolean


Public Function Ok As Boolean
  Dim SQL As Object
  Set SQL = NewQuery

  Dim S As Object
  Set S = NewQuery
  S.Add("SELECT CONTROLEDEACESSO FROM SAM_PARAMETROSPRESTADOR")
  S.Active = True

  'GArcia
  'If S.FieldByName("CONTROLEDEACESSO").Value = "N" Then
  '  Ok = True
  '  Set S=Nothing
  '  Exit Function
  'End If

  SQL.Add("Select SAM_PRESTADOR_PROC.DATAINICIAL, SAM_PRESTADOR_PROC.DATAFINAL,SAM_PRESTADOR_PROC.RESPONSAVEL,SAM_PRESTADOR.filialpadrao FROM SAM_PRESTADOR_PROC, sam_prestador WHERE SAM_PRESTADOR_PROC.Handle = :HANDLE And  SAM_PRESTADOR.Handle = SAM_PRESTADOR_PROC.prestador")
  SQL.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_PROC")

  SQL.Active = True

   Ok = IIf(SQL.FieldByName("DATAFINAL").IsNull And ((SQL.FieldByName("RESPONSAVEL").AsInteger = CurrentUser) Or (SQL.FieldByName("FILIALPADRAO").IsNull)), True, False)

  'SQL.Add("SELECT DATAINICIAL,DATAFINAL,RESPONSAVEL FROM SAM_PRESTADOR_PROC WHERE HANDLE = :HANDLE")
  'SQL.ParamByName("HANDLE").Value=RecordHandleOfTable("SAM_PRESTADOR_PROC")
  'SQL.Active=True
  vDataInicial = SQL.FieldByName("DATAINICIAL").AsDateTime
  'Ok = IIf(SQL.FieldByName("DATAFINAL").IsNull And SQL.FieldByName("RESPONSAVEL").AsInteger = CurrentUser,True,False)
  If Not SQL.FieldByName("DATAFINAL").IsNull Then
    Mensagem = "Processo finalizado! Operação não permitida." + Chr(13)
  End If
  If SQL.FieldByName("RESPONSAVEL").AsInteger <> CurrentUser Then
    Mensagem = Mensagem + "Usuário não é o responsável!"
  End If
  Set SQL = Nothing
End Function

Public Sub BOTAOCOBRARRATIFICACAO_OnClick()

	If (CurrentQuery.State = 2 Or CurrentQuery.State = 3)  Then
	  bsShowMessage("Ação não permitida. A fase está em edição.","I")
	  Exit Sub
	End If

	Dim componente As CSBusinessComponent
	Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcFasesBLL, Benner.Saude.Prestadores.Business")
	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
	componente.Execute("CobrarRatificacao")

	Set componente = Nothing
	Exit Sub
End Sub

Public Sub BOTAOENVIARTERMO_OnClick()

  If (CurrentQuery.State = 2 Or CurrentQuery.State = 3)  Then
  	bsShowMessage("Ação não permitida. A fase está em edição.","I")
  	Exit Sub
  End If

  If Not VerificarProcessoFinalizado Then

  	If(CurrentQuery.FieldByName("ARQUIVOTERMO").IsNull) Then
  	  bsShowMessage("Nenhum arquivo indicado para envio do termo!","I")
  	  Exit Sub
    Else
  	On Error GoTo Erro

	  Dim TvFormEnviarEmailBLL As CSBusinessComponent

	  Set TvFormEnviarEmailBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.TvFormEnviarEmailBLL, Benner.Saude.Prestadores.Business")
	  TvFormEnviarEmailBLL.AddParameter(pdtInteger, 2)
	  TvFormEnviarEmailBLL.AddParameter(pdtString, "")
	  TvFormEnviarEmailBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
	  TvFormEnviarEmailBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("PRESTADOR").AsInteger)
	  TvFormEnviarEmailBLL.AddParameter(pdtInteger, 0)
	  TvFormEnviarEmailBLL.AddParameter(pdtAutomatic, False)
	  TvFormEnviarEmailBLL.Execute("PreencherFormularioEnvioEmail")
	  Set TvFormEnviarEmailBLL = Nothing
	  RefreshNodesWithTable("")
	  Exit Sub

	Erro:
	  bsShowMessage(Err.Description, "E")
	  Set TvFormEnviarEmailBLL = Nothing
      Exit Sub
	End If

  End If

End Sub

Public Sub BOTAOGERARTEXTODOU_OnClick()

  If (CurrentQuery.State = 2 Or CurrentQuery.State = 3)  Then
	  bsShowMessage("Ação não permitida. A fase está em edição.","I")
	  Exit Sub
	End If

  If Not VerificarProcessoFinalizado Then

    podeContinuar = True

	On Error GoTo Erro

	  Dim SamPrestadorProcFasesBLL As CSBusinessComponent

	  Set SamPrestadorProcFasesBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcFasesBLL, Benner.Saude.Prestadores.Business")
	  SamPrestadorProcFasesBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
	  SamPrestadorProcFasesBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("PRESTADOR").AsInteger)
	  SamPrestadorProcFasesBLL.Execute("PreencherCampoTextoPublicacao")

	  bsShowMessage("Texto para D.O.U. gerado com sucesso.", "I")

	  Set SamPrestadorProcFasesBLL = Nothing
	  RefreshNodesWithTable("")
	  Exit Sub

    Erro:
	  bsShowMessage(Err.Description, "E")
	  podeContinuar = False
      Set SamPrestadorProcFasesBLL = Nothing
    Exit Sub

  End If

End Sub

Public Function VerificarProcessoFinalizado As Boolean

	Dim SamPrestadorProcFasesBLL As CSBusinessComponent
	Set SamPrestadorProcFasesBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcFasesBLL, Benner.Saude.Prestadores.Business")

	SamPrestadorProcFasesBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
    VerificarProcessoFinalizado = SamPrestadorProcFasesBLL.Execute("VerificarCredenciamentoProcessado")

    If VerificarProcessoFinalizado Then
		bsShowMessage("Processo finalizado! Operação não permitida.", "I")
	End If

	Set SamPrestadorProcFasesBLL = Nothing

End Function

Public Sub BOTAOPREENCHERTERMO_OnClick()
  On Error GoTo Err

  podeContinuar = True

  If ValidacoesPreencherTermo Then

    If(CurrentQuery.FieldByName("PARECER").Value = "A" Or CurrentQuery.FieldByName("PARECER").Value = "V") Then

	  'FORMATAR O ARQUIVO DE ACORDO COM O TIPO DE PROCESSO DA FASE SENDO INSERIDA
	  Dim SamPrestadorProcFasesBLL As CSBusinessComponent
	  Set SamPrestadorProcFasesBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcFasesBLL, Benner.Saude.Prestadores.Business")

	  SamPrestadorProcFasesBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("TIPOCREDENCTO").AsInteger)
	  SamPrestadorProcFasesBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
	  SamPrestadorProcFasesBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("PRESTADOR").AsInteger)
	  SamPrestadorProcFasesBLL.AddParameter(pdtAutomatic, executadoPelaWeb )
	  SamPrestadorProcFasesBLL.Execute("ModificarModeloTermo")

	  Set SamPrestadorProcFasesBLL = Nothing
	  RefreshNodesWithTable("")
    End If
  End If
  Exit Sub

    Err:
        bsShowMessage(Err.Description, "E")
        podeContinuar = False
    Exit Sub
End Sub

Public Function ValidacoesPreencherTermo As Boolean

  If (CurrentQuery.State = 2 Or CurrentQuery.State = 3)  Then
	  bsShowMessage("Ação não permitida. A fase está em edição.","I")
	  ValidacoesPreencherTermo = False
	  Exit Function
  End If

  If(CurrentQuery.FieldByName("ENVIOSTERMO").Value > 0) Then
  	bsShowMessage("Documento não pode ser substituído pois já foi enviado para o prestador.", "I")
  	ValidacoesPreencherTermo = False
  	Exit Function
  End If

  If Not (CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
    bsShowMessage("Ação não permitida. A fase foi finalizada.", "I")
    ValidacoesPreencherTermo = False
    Exit Function
  End If

  ValidacoesPreencherTermo = True
End Function


Public Sub BOTAORATIFICAR_OnClick()

  If (CurrentQuery.State = 2 Or CurrentQuery.State = 3)  Then
    bsShowMessage("Ação não permitida. A fase está em edição.","I")
	Exit Sub
  End If

  RatificarFase("R")
End Sub

Public Sub BOTAORATIFICARPOROFICIO_OnClick()

  If (CurrentQuery.State = 2 Or CurrentQuery.State = 3)  Then
    bsShowMessage("Ação não permitida. A fase está em edição.","I")
	Exit Sub
  End If

  RatificarFase("O")
End Sub



Public Sub RatificarFase(pTipoRatificacao As String)
	On Error GoTo Exception

	Dim componente As CSBusinessComponent
	Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcFasesBLL, Benner.Saude.Prestadores.Business")
	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
	componente.AddParameter(pdtString, pTipoRatificacao)
	componente.Execute("RatificarFase")

	Set componente = Nothing
	RefreshNodesWithTable("")
	Exit Sub

	Exception:
    	Set componente = Nothing
    	bsShowMessage(Err.Description, "I")
    	Exit Sub
End Sub

Public Sub ConfirmarRatificacao(pTipoRatificacao As String)
	Dim componente As CSBusinessComponent

	Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcFasesBLL, Benner.Saude.Prestadores.Business")
	componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
	componente.AddParameter(pdtString, pTipoRatificacao)
	componente.Execute("ConfirmarRatificacaoFase")
	Set componente = Nothing

	RefreshNodesWithTable("")
	Exit Sub
End Sub


Public Sub TABLE_AfterInsert()
  If Not Ok Then
    RefreshNodesWithTable "SAM_PRESTADOR_PROC"
    bsShowMessage(Mensagem, "E")
    CurrentQuery.Cancel
    RefreshNodesWithTable "SAM_PRESTADOR_PROC_FASES"
  End If

End Sub

Public Sub TABLE_AfterPost()
  VerificaDataFinal
End Sub

Public Sub TABLE_AfterScroll()

  Dim componente As CSBusinessComponent
  Dim vControlaRatificacao As Boolean

  Set componente = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcFasesBLL, Benner.Saude.Prestadores.Business")
  componente.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
  vControlaRatificacao = componente.Execute("VerificarTipoFaseControleRatificacao")

  BOTAORATIFICAR.Visible = vControlaRatificacao
  BOTAORATIFICARPOROFICIO.Visible = vControlaRatificacao
  BOTAOCOBRARRATIFICACAO.Visible = vControlaRatificacao

  Set componente = Nothing

  Dim SamPrestadorProcBLL As CSBusinessComponent
  Dim visivel As Boolean

  Set SamPrestadorProcBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcBLL, Benner.Saude.Prestadores.Business")
  SamPrestadorProcBLL.AddParameter(pdtString, "CREDENCIAMENTOAVANCADO")
  visivel = SamPrestadorProcBLL.Execute("VerificarParametrosParaCredenciamentoAutomatico")
  BOTAOENVIARTERMO.Visible = visivel
  CONTROLESAVANCADOS.Visible = visivel
  Set SamPrestadorProcBLL = Nothing

  Dim SamPrestadorProcFasesBLL As CSBusinessComponent
  Set SamPrestadorProcFasesBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcFasesBLL, Benner.Saude.Prestadores.Business")

  BotoesVisiveis SamPrestadorProcFasesBLL, "DOU"
  BotoesVisiveis SamPrestadorProcFasesBLL, "TERMO"


  Set SamPrestadorProcFasesBLL = Nothing

    If(Not CurrentQuery.FieldByName("ENVIOSTERMO").IsNull And CurrentQuery.FieldByName("ENVIOSTERMO").Value > 0 ) Then
		ARQUIVOTERMO.ReadOnly = True
	    PERCENTUALREAJUSTE.ReadOnly = True
	Else
	    ARQUIVOTERMO.ReadOnly = False
	    PERCENTUALREAJUSTE.ReadOnly = False

  End If

  VerificaDataFinal
  executadoPelaWeb = False
End Sub

Public Sub BotoesVisiveis(SamPrestadorProcFasesBLL As CSBusinessComponent, valor As String)

  SamPrestadorProcFasesBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("PRESTADOR").AsInteger)
  SamPrestadorProcFasesBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("TIPOCREDENCIAMETOFASE").AsInteger)

  If(valor = "DOU") Then
    SamPrestadorProcFasesBLL.AddParameter(pdtString, valor)
	BOTAOGERARTEXTODOU.Visible = SamPrestadorProcFasesBLL.Execute("VerificarDouTermo")
  Else
    SamPrestadorProcFasesBLL.AddParameter(pdtString, valor)
	BOTAOPREENCHERTERMO.Visible = SamPrestadorProcFasesBLL.Execute("VerificarDouTermo")
  End If

  SamPrestadorProcFasesBLL.ClearParameters()
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vDataI, vDataF, Linha As String
  If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime < vDataInicial Then
    CanContinue = False
    bsShowMessage("Data inicial da fase não pode ser anterior à data inicial do processo.", "E")
  End If
  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    If CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAFINAL").AsDateTime Then
      CanContinue = False
      bsShowMessage("Data inicial da fase não pode ser maior que a data final", "E")

      Exit Sub
    End If

    If CurrentQuery.FieldByName("DATAFINAL").AsDateTime > ServerDate Then
      CanContinue = False
      bsShowMessage("Data final da fase não pode ser maior que a data de hoje", "E")

      Exit Sub
    End If


  End If
  'VERIRIFICAR SE A FASE ATUAL É ULTIMA FASE
  'SE = ULTIMA FASE ENTAO NAO PERMITIR DUAS IGUAIS
  Dim SQL As Object
  Set SQL = NewQuery
  SQL.Add("SELECT * FROM SAM_TIPOCREDENCIAMENTO_FASE WHERE HANDLE = :FASE")
  SQL.ParamByName("FASE").Value = CurrentQuery.FieldByName("TIPOCREDENCIAMETOFASE").AsInteger
  SQL.Active = True
  If SQL.FieldByName("ULTIMAFASE").AsString = "S" Then
    SQL.Clear
    SQL.Add("SELECT * FROM SAM_PRESTADOR_PROC_FASES WHERE PRESTADORPROCESSO = :PRESTADORPROC AND")
    SQL.Add("TIPOCREDENCIAMETOFASE = :FASE AND HANDLE <> :HANDLE")
    SQL.ParamByName("PRESTADORPROC").Value = CurrentQuery.FieldByName("PRESTADORPROCESSO").AsInteger
    SQL.ParamByName("FASE").Value = CurrentQuery.FieldByName("TIPOCREDENCIAMETOFASE").AsInteger
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True
    If Not SQL.EOF Then
      CanContinue = False
      bsShowMessage("Última fase já cadastrada!", "E")
      Set SQL = Nothing
      Exit Sub
    End If
  Else
    SQL.Clear
    SQL.Add("SELECT * FROM SAM_PRESTADOR_PROC_FASES WHERE PRESTADORPROCESSO = :PRESTADORPROC AND")
    SQL.Add("TIPOCREDENCIAMETOFASE = :FASE AND RESPONSAVEL = :RESPONSAVEL AND HANDLE <> :HANDLE")
    SQL.ParamByName("PRESTADORPROC").Value = CurrentQuery.FieldByName("PRESTADORPROCESSO").AsInteger
    SQL.ParamByName("RESPONSAVEL").Value = CurrentQuery.FieldByName("RESPONSAVEL").AsInteger
    SQL.ParamByName("FASE").Value = CurrentQuery.FieldByName("TIPOCREDENCIAMETOFASE").AsInteger
    SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.Active = True
    If Not SQL.EOF Then
      CanContinue = False
      bsShowMessage("Fase já cadastrada por este responsável!", "E")
      Set SQL = Nothing
      Exit Sub
    End If
  End If
  SQL.Active = False
  Set SQL = Nothing

  'NÃO PERMITIR FINALIZAR UMA FASE QUE ESTEJA EM ANALISE
  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    If CurrentQuery.FieldByName("PARECER").AsString = "A" Then
      CanContinue = False
      bsShowMessage("Fase está em análise, impossibilitado de efetuar a conclusão da fase!", "E")
      Exit Sub
    End If
  End If

  If (CurrentQuery.FieldByName("PARECER").AsString <> "I") And (Not CurrentQuery.FieldByName("JUSTIFICATIVA").IsNull) Then
      CanContinue = False
      bsShowMessage("Justificativa do Indeferimento só deve ser informada se o parecer for 'Indeferido'.", "E")
      Exit Sub
  End If


  If (Not CurrentQuery.FieldByName("DATAPUBLICACAODOU").IsNull And Not CurrentQuery.FieldByName("TEXTODOU").IsNull)  Then
  	TEXTODOU.ReadOnly = True
  Else
	TEXTODOU.ReadOnly = False
  End If


  On Error GoTo Erro

 'REGRAS PARA QUANDO A FASE ESTIVER "DEFERIDO"
  Dim SamPrestadorProcBLL As CSBusinessComponent

  Set SamPrestadorProcBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcBLL, Benner.Saude.Prestadores.Business")
  SamPrestadorProcBLL.AddParameter(pdtString, "CREDENCIAMENTOAVANCADO")

  If( (CurrentQuery.FieldByName("PARECER").AsString = "D") And (SamPrestadorProcBLL.Execute("VerificarParametrosParaCredenciamentoAutomatico")) )Then

	Dim SamPrestadorProcFasesBLL As CSBusinessComponent
	Set SamPrestadorProcFasesBLL = BusinessComponent.CreateInstance("Benner.Saude.Prestadores.Business.ProcessoCredDescred.SamPrestadorProcFasesBLL, Benner.Saude.Prestadores.Business")

	SamPrestadorProcFasesBLL.AddParameter(pdtDateTime, CurrentQuery.FieldByName("DATAPUBLICACAODOU").AsDateTime)
	SamPrestadorProcFasesBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("PRESTADORPROCESSO").AsInteger)
	SamPrestadorProcFasesBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("TIPOCREDENCIAMETOFASE").AsInteger)
	SamPrestadorProcFasesBLL.AddParameter(pdtAutomatic, Not CurrentQuery.FieldByName("DATARATIFICACAO").IsNull)
	SamPrestadorProcFasesBLL.AddParameter(pdtAutomatic, Not CurrentQuery.FieldByName("DATAFINAL").IsNull)

	SamPrestadorProcFasesBLL.Execute("VerificarConfiguracoesFaseDeferido")
    Set SamPrestadorProcBLL = Nothing
    Set SamPrestadorProcFasesBLL = Nothing


    If (CurrentQuery.FieldByName("DATAFINAL").IsNull) Then
	  CanContinue = False
	  bsShowMessage("A fase se encontra aberta.", "E")
	  Exit Sub
    End If
  Exit Sub

	Erro:
		bsShowMessage(Err.Description, "E")
		CanContinue = False
	Exit Sub
  End If

  Dim terminaCom As String
  If(Not CurrentQuery.FieldByName("ARQUIVOTERMO").IsNull) Then
	 terminaCom = LCase(CurrentQuery.FieldByName("ARQUIVOTERMO").Value)
	 If(terminaCom Like "*.rtf") Then
       CanContinue = True
     Else
       bsShowMessage("Arquivo de Modelo de Termo de Credenciamento deve ter o formato/extensão RTF!", "E")
       CanContinue = False
     End If
  End If


End Sub

Function VerificaDataFinal
  If CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    DATAINICIAL.ReadOnly = False
    DATAFINAL.ReadOnly = False
    PARECER.ReadOnly = False
    RESPONSAVEL.ReadOnly = False
    TIPOCREDENCIAMETOFASE.ReadOnly = False
    TIPOCREDENCTO.ReadOnly = False
  Else
    DATAINICIAL.ReadOnly = True
    DATAFINAL.ReadOnly = True
    PARECER.ReadOnly = True
    RESPONSAVEL.ReadOnly = True
    TIPOCREDENCIAMETOFASE.ReadOnly = True
    TIPOCREDENCTO.ReadOnly = True
  End If
End Function

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

  Dim Msg As String

  If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok

  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
    CanContinue = False
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("ENVIOSTERMO").Value > 0) Then
    bsShowMessage("Houve emissão de Termo de Credenciamento. Impossível excluir a fase.", "E")
  	CanContinue = False
  End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  Dim Msg As String

  If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  CanContinue = Ok

  If Not CanContinue Then
    bsShowMessage(Mensagem, "E")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    bsShowMessage("Fase com data finalizada não pode ser alterada!", "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  Dim Msg As String
  If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then
    bsShowMessage(Msg, "E")
    CanContinue = False
    Exit Sub
  End If

  If Not Ok Then
    bsShowMessage(Mensagem, "E")
    CanContinue = False
    Exit Sub
  End If

End Sub

Public Sub TABLE_NewRecord()
  CurrentQuery.FieldByName("RESPONSAVEL").Value = CurrentUser

  If WebMode Then
    CurrentQuery.FieldByName("PRESTADOR").AsInteger = RecordHandleOfTable("SAM_PRESTADOR")

    Dim SQL As Object
    Set SQL = NewQuery

    SQL.Add("SELECT TIPOCREDENCIAMENTO")
    SQL.Add("FROM SAM_PRESTADOR_PROC_CREDEN")
    SQL.Add("WHERE HANDLE = :HPROCCREDEN")
    SQL.ParamByName("HPROCCREDEN").AsInteger = RecordHandleOfTable("SAM_PRESTADOR_PROC_CREDEN")
    SQL.Active = True

    CurrentQuery.FieldByName("TIPOCREDENCTO").AsInteger = SQL.FieldByName("TIPOCREDENCIAMENTO").AsInteger
  Else
	Dim qPrestador As Object
	Set qPrestador = NewQuery
	qPrestador.Clear

	qPrestador.Add("SELECT PRESTADOR FROM SAM_PRESTADOR_PROC WHERE HANDLE = :PHANDLE")
    qPrestador.ParamByName("PHANDLE").AsInteger = RecordHandleOfTable("SAM_PRESTADOR_PROC")
	qPrestador.Active = True

	CurrentQuery.FieldByName("PRESTADOR").AsInteger = qPrestador.FieldByName("PRESTADOR").AsInteger

	Set qPrestador = Nothing

  End If
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  Dim mensagemErro As String

	If (CommandID = "BOTAOGERARTEXTODOU") Then
		 BOTAOGERARTEXTODOU_OnClick
		 CanContinue = podeContinuar
    End If

    If (CommandID = "BOTAOPREENCHERTERMO") Then
         executadoPelaWeb = True
		 BOTAOPREENCHERTERMO_OnClick
		 CanContinue = podeContinuar
    End If

    If (CommandID = "BOTAORATIFICAR") Then
    	ConfirmarRatificacao("R")
    End If

    If (CommandID = "BOTAORATIFICARPOROFICIO") Then
		ConfirmarRatificacao("O")
    End If

End Sub

