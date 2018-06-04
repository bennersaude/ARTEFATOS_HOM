'HASH: EBA679D5FC8B92169BD31570EC9D97BE
'#uses "*CriaTabelaTemporariaSqlServer"
'#uses "*bsShowMessage"
Option Explicit

Dim gsDadosEventoSolicitadoAntes As String
Dim gsDadosEventoSolicitadoDepois As String

Public Sub TABLE_AfterEdit()
  Dim DLL As Object
  Set DLL = CreateBennerObject("CA043.Autorizacao")
  gsDadosEventoSolicitadoAntes = DLL.ObterDadosEventoSolicitado(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set DLL = Nothing
End Sub

Public Sub BOTAOPFINTEGRAL_OnClick()
  revalidar
End Sub

Public Sub TABLE_AfterPost()
  Dim viRetorno As Integer
  Dim vsMensagemErro As String
  Dim DLL As Object

  Set DLL = CreateBennerObject("CA043.Autorizacao")
  gsDadosEventoSolicitadoDepois = DLL.ObterDadosEventoSolicitado(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  viRetorno = DLL.EventoSolicitadoConsistenciasAposAlterar(CurrentSystem, gsDadosEventoSolicitadoAntes, gsDadosEventoSolicitadoDepois, vsMensagemErro)
  Set DLL = Nothing
  If viRetorno > 0 Then
    bsShowMessage(vsMensagemErro, "E")
  End If

  'revalidar o evento
  revalidar
End Sub

Public Sub TABLE_AfterScroll()
  BOTAOALERTA.Enabled = False
  BOTAOALTERAREQUIPEVIA.Enabled = False
  BOTAOAUTORIZAR.Enabled = False
  BOTAOCANCELAR.Enabled = False
  BOTAOLIVREESCOLHA.Enabled = False
  BOTAOPFINTEGRAL.Enabled = True
  BOTAOPRORROGARINTERNACAO.Enabled = False
  BOTAOREVERTER.Enabled = False
  BOTAOREVERTEROUTROUSUARIO.Enabled = False
  CODIGOTABELA.WebLocalWhere = " A.HANDLE IN (SELECT TABELATISS FROM SAM_TGE_TABELATISS WHERE EVENTO = @~CAMPO(EVENTO))"

  SessionVar("HANDLEEVENTOSOLICIT") = CurrentQuery.FieldByName("HANDLE").AsString

  Dim qTipoAutorizacao As Object
  Set qTipoAutorizacao = NewQuery

  qTipoAutorizacao.Add("SELECT T.PERMITEALTERARRECEBEDOR")
  qTipoAutorizacao.Add("  FROM SAM_AUTORIZ A")
  qTipoAutorizacao.Add("  JOIN SAM_TIPOAUTORIZ T ON (T.HANDLE = A.TIPOAUTORIZACAO)")
  qTipoAutorizacao.Add(" WHERE A.HANDLE = :HANDLE")
  qTipoAutorizacao.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
  qTipoAutorizacao.Active = True
  If qTipoAutorizacao.FieldByName("PERMITEALTERARRECEBEDOR").AsString = "S" Then
    RECEBEDOR.ReadOnly = False
  Else
    RECEBEDOR.ReadOnly = True
  End If

  qTipoAutorizacao.Active = False
  Set qTipoAutorizacao = Nothing

  RecordReadOnly = True

  ' informações
  ROTULO1.Text = pegaRotulo1
  If VisibleMode Then
    ROTULO2.Visible = False
  Else
    ROTULO2.Text = pegaRotulo2
  End If
End Sub

Public Function pegaRotulo1
	Dim sql As BPesquisa
	Set sql=NewQuery
	sql.Add("SELECT COUNT(1) NREC FROM SAM_AUTORIZ_EVENTOGERADO WHERE EVENTOSOLICITADO=:ES AND SITUACAO IN ('C', 'N') ")
	sql.ParamByName("ES").AsInteger=CurrentQuery.FieldByName("HANDLE").AsInteger
	sql.Active=True
	If sql.FieldByName("NREC").AsInteger>0 Then
	    If VisibleMode Then
			pegaRotulo1 = "Esta solicitação possui eventos negados"
	    Else
			pegaRotulo1 = "@ <p class=""frmerror""> <IMG SRC=""img/alert.gif"" /> Esta solicitação possui eventos negados</p>"
		End If
	Else
        sql.Clear
        sql.Add("SELECT COUNT(1) NREC FROM SAM_AUTORIZ_EVENTOGERADO WHERE EVENTOSOLICITADO=:ES")
	    sql.ParamByName("ES").AsInteger=CurrentQuery.FieldByName("HANDLE").AsInteger
	    sql.Active=True
	    If sql.FieldByName("NREC").AsInteger>0 Then
	      If VisibleMode Then
            pegaRotulo1 = "Autorizado"
	      Else
			pegaRotulo1 = "@ <p class=""frmerror""> <IMG SRC=""img/link.gif"" /> Autorizado</p>"
		  End If
		Else
		  If VisibleMode Then
            pegaRotulo1 = "Não existem Eventos Gerados para esta Solicitação"
		  Else
		    pegaRotulo1 = "@ <p class=""frmerror""> <IMG SRC=""img/link.gif"" /> Não existem Eventos Gerados para esta Solicitação</p>"
		  End If
		End If
	End If
	Set sql=Nothing
End Function

Public Function pegaRotulo2
	Dim alertas As String
	alertas = SessionVar("alertas")
	If alertas <> "" Then
		Dim DLL As Object
		Dim retorno As Integer
		Dim alertasFormatados As String
		Dim mensagem As String
		Set DLL = CreateBennerObject("samauto.autorizador")
		retorno = DLL.mostrarAlertas(CurrentSystem, alertas, alertasFormatados, mensagem)
		If retorno > 0 Then
			pegaRotulo2 = mensagem
		Else
			pegaRotulo2 = "@"+alertasFormatados
		End If
		SessionVar("alertas") = ""
	End If
End Function


Public Sub validar
	Dim sql As Object
    Set sql = NewQuery

	sql.Active = False
	sql.Clear
	sql.Add("SELECT HANDLE, WEBAUTORIZ FROM WEB_AUTORIZ_EVENTOS WHERE EVENTOSOLICITADO=:ES")
	sql.ParamByName("ES").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	sql.Active = True

	If sql.FieldByName("HANDLE").AsInteger > 0 Then
		Dim SP2 As Object
		Set SP2 = NewStoredProc
		SP2.Name = "BSAUT_AUTORIZINSEREEVENTOSWEB"
		SP2.AddParam("P_WEBAUTORIZ", ptInput)
		SP2.ParamByName("P_WEBAUTORIZ").DataType = ftInteger
		SP2.AddParam("P_WEBAUTORIZEVENTO", ptInput)
		SP2.ParamByName("P_WEBAUTORIZEVENTO").DataType = ftInteger
		SP2.AddParam("P_HANDLEAUTORIZ", ptInput)
		SP2.ParamByName("P_HANDLEAUTORIZ").DataType = ftInteger
		SP2.AddParam("P_USUARIO", ptInput)
		SP2.ParamByName("P_USUARIO").DataType = ftInteger
		SP2.AddParam("P_TIPOOPERACAOTISS", ptInput)
		SP2.ParamByName("P_TIPOOPERACAOTISS").DataType = ftString


		SP2.ParamByName("P_WEBAUTORIZ").AsInteger = sql.FieldByName("WEBAUTORIZ").AsInteger
		SP2.ParamByName("P_WEBAUTORIZEVENTO").AsInteger = sql.FieldByName("HANDLE").AsInteger
		SP2.ParamByName("P_HANDLEAUTORIZ").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
		SP2.ParamByName("P_USUARIO").AsInteger = CurrentUser
		SP2.ParamByName("P_TIPOOPERACAOTISS").AsString = "S"
		SP2.ExecProc

		Set SP2 = Nothing
		InfoDescription = "O evento foi validado, pois não haviam eventos gerados."

	End If


   	Set sql = Nothing
End Sub


Public Sub revalidar
	CriaTabelaTemporariaSqlServer
	Dim retorno As Long

	Dim sql As BPesquisa
	Set sql = NewQuery
	sql.Add("SELECT COUNT(1) NREC FROM SAM_AUTORIZ_EVENTOGERADO WHERE EVENTOSOLICITADO=:ES")
	sql.ParamByName("ES").AsInteger=CurrentQuery.FieldByName("HANDLE").AsInteger
	sql.Active=True
	If sql.FieldByName("NREC").AsInteger = 0 Then
	  validar
	Else

		Dim SQLAgendado As Object
		Set SQLAgendado = NewQuery

		SQLAgendado.Add("SELECT EXECUTAAUTORIZACAOAGENDADA FROM SAM_PARAMETROSWEB")
		SQLAgendado.Active = True

		If SQLAgendado.FieldByName("EXECUTAAUTORIZACAOAGENDADA").AsString = "S" Then
			Dim vsMensagemErro As String
			Dim Obj As Object
			Dim vcContainer As CSDContainer
			Set vcContainer = NewContainer
			vcContainer.AddFields("HANDLE:INTEGER")

			vcContainer.Insert
			vcContainer.Field("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

			Set Obj = CreateBennerObject("BSServerExec.ProcessosServidor")
			retorno = Obj.ExecucaoImediata(CurrentSystem, _
										  "CA043", _
										  "RevalidarSolicitado", _
										  "Processamento de Autorização", _
										  CurrentQuery.FieldByName("HANDLE").AsInteger, _
										  "SAM_AUTORIZ_EVENTOSOLICIT", _
										  "SITUACAOPROCESSAMENTO", _
										  "", _
										  "", _
										  "P", _
										  True, _
										  vsMensagemErro, _
										  vcContainer)
			If retorno = 0 Then
				bsShowMessage("Processo enviado para execução no servidor!", "I")
				bsShowMessage("Autorização sendo revalidada!", "I")

			Else
				bsShowMessage("Erro ao enviar processo para execução no servidor!" + Chr(13) + vsMensagemErro, "I")
			End If

		Else
			Dim mensagem As String
			Dim alertas As String

			Dim dll As Object
			Set dll=CreateBennerObject("ca043.autorizacao")
			retorno = dll.revalidarSolicitado(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, alertas, mensagem)
			Set dll=Nothing
			If retorno > 0 Then
				bsShowMessage(mensagem, "I")
			Else
				SessionVar("alertas") = alertas
				bsShowMessage("Revalidação concluída com sucesso", "I")
			End If
		End If
	End If
	Set sql=Nothing

End Sub


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If VisibleMode Then
    bsShowMessage("Operação permitida apenas pela Interface de Autorização!", "E")
    CanContinue = False
  Else
    Dim qSQL As Object
    Set qSQL = NewQuery

    qSQL.Clear
    qSQL.Add("SELECT COUNT(*) QTDREGISTROS")
    qSQL.Add("FROM SAM_AUTORIZ_ANOTADM")
    qSQL.Add("WHERE EVENTOSOLICITADO = :HEVENTOSOLICITADO")
    qSQL.ParamByName("HEVENTOSOLICITADO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qSQL.Active = True

    If qSQL.FieldByName("QTDREGISTROS").AsInteger > 0 Then
      bsShowMessage("Evento possui anotações vinculadas! Operação cancelada.", "E")
      CanContinue = False
      Set qSQL = Nothing
      Exit Sub
    End If

    qSQL.Clear
    qSQL.Add("SELECT COUNT(*) QTDREGISTROS")
    qSQL.Add("FROM SAM_AUTORIZ_EVENTOGERADO")
    qSQL.Add("WHERE EVENTOSOLICITADO = :HEVENTOSOLICITADO")
    qSQL.ParamByName("HEVENTOSOLICITADO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qSQL.Active = True

    If qSQL.FieldByName("QTDREGISTROS").AsInteger > 0 Then
      bsShowMessage("Evento possui eventos gerados! Operação cancelada.", "E")
      CanContinue = False
      Set qSQL = Nothing
      Exit Sub
    End If

    qSQL.Clear
    qSQL.Add("SELECT COUNT(*) QTDREGISTROS")
    qSQL.Add("FROM SAM_AUTORIZ_EVENTOSOLICIT_DOC")
    qSQL.Add("WHERE EVENTOSOLICITADO = :HEVENTOSOLICITADO")
    qSQL.ParamByName("HEVENTOSOLICITADO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qSQL.Active = True

    If qSQL.FieldByName("QTDREGISTROS").AsInteger > 0 Then
      bsShowMessage("Evento possui documentos vinculados! Operação cancelada.", "E")
      CanContinue = False
      Set qSQL = Nothing
      Exit Sub
    End If

    qSQL.Clear
    qSQL.Add("SELECT COUNT(*) QTDREGISTROS")
    qSQL.Add("FROM SAM_AUTORIZ_EVENTOSOLICITCID")
    qSQL.Add("WHERE EVENTOSOLICITADO = :HEVENTOSOLICITADO")
    qSQL.ParamByName("HEVENTOSOLICITADO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qSQL.Active = True

    If qSQL.FieldByName("QTDREGISTROS").AsInteger > 0 Then
      bsShowMessage("Evento possui CID's vinculados! Operação cancelada.", "E")
      CanContinue = False
      Set qSQL = Nothing
      Exit Sub
    End If

    qSQL.Clear
    qSQL.Add("SELECT COUNT(*) QTDREGISTROS")
    qSQL.Add("FROM SAM_AUTORIZ_MATMED")
    qSQL.Add("WHERE EVENTOSOLICITADO = :HEVENTOSOLICITADO")
    qSQL.ParamByName("HEVENTOSOLICITADO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qSQL.Active = True

    If qSQL.FieldByName("QTDREGISTROS").AsInteger > 0 Then
      bsShowMessage("Evento possui Materiais/Medicamentos vinculados! Operação cancelada.", "E")
      CanContinue = False
      Set qSQL = Nothing
      Exit Sub
    End If

    qSQL.Clear
    qSQL.Add("SELECT COUNT(*) QTDREGISTROS")
    qSQL.Add("FROM SAM_AUTORIZ_OBSERVACAO")
    qSQL.Add("WHERE EVENTOSOLICIT = :HEVENTOSOLICITADO")
    qSQL.ParamByName("HEVENTOSOLICITADO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qSQL.Active = True

    If qSQL.FieldByName("QTDREGISTROS").AsInteger > 0 Then
      bsShowMessage("Evento possui observações vinculadas! Operação cancelada.", "E")
      CanContinue = False
      Set qSQL = Nothing
      Exit Sub
    End If

    qSQL.Clear
    qSQL.Add("SELECT COUNT(*) QTDREGISTROS")
    qSQL.Add("FROM SAM_AUTORIZ_RECOMENDACAO")
    qSQL.Add("WHERE EVENTOSOLICITADO = :HEVENTOSOLICITADO")
    qSQL.ParamByName("HEVENTOSOLICITADO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    qSQL.Active = True

    If qSQL.FieldByName("QTDREGISTROS").AsInteger > 0 Then
      bsShowMessage("Evento possui recomendações vinculadas! Operação cancelada.", "E")
      CanContinue = False
      Set qSQL = Nothing
      Exit Sub
    End If

    Set qSQL = Nothing
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)

  If VisibleMode Then
    bsShowMessage("Operação permitida apenas pela Interface de Autorização!", "E")
    CanContinue = False
  Else

    RecordReadOnly = True

    Dim qSQL As Object
    Set qSQL = NewQuery

    qSQL.Add("SELECT SITUACAO")
    qSQL.Add("FROM SAM_AUTORIZ")
    qSQL.Add("WHERE HANDLE = :HAUTORIZ")
    qSQL.ParamByName("HAUTORIZ").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
    qSQL.Active = True

    If qSQL.FieldByName("SITUACAO").AsString <> "A" Then
      bsShowMessage("Somente autorizações ""Abertas"" permitem alteração!", "E")
      CanContinue = False
    End If

    Set qSQL = Nothing
  End If


End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qVerificaJaPago As Object
  Set qVerificaJaPago = NewQuery

  qVerificaJaPago.Clear
  qVerificaJaPago.Add("SELECT SUM(QTDPAGA) QTD")
  qVerificaJaPago.Add("  FROM SAM_AUTORIZ_EVENTOGERADO")
  qVerificaJaPago.Add(" WHERE EVENTOSOLICITADO = :EVENTOSOLICITADO")
  qVerificaJaPago.ParamByName("EVENTOSOLICITADO").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qVerificaJaPago.Active = True
  If qVerificaJaPago.FieldByName("QTD").AsInteger > 0 Then
    'devemos deixar alterar agora somente se a quantidade colocada for maior que a qtd já paga
    If qVerificaJaPago.FieldByName("QTD").AsInteger > CurrentQuery.FieldByName("QTDSOLICITADA").AsInteger Then
      bsShowMessage("Alteração não permitida. Quantidade solicitada menor que quantidade já paga para o evento", "E")
      CanContinue = False
    End If
  End If

  Dim dll As Object
  Dim mensagem As String
  Dim retorno As Boolean

  Set dll = CreateBennerObject("ca043.autorizacao")

  retorno = dll.VerificaDiariaFrqInternacaoFaturada(CurrentSystem, CurrentQuery.FieldByName("AUTORIZACAO").AsInteger, CurrentQuery.FieldByName("HANDLE").AsInteger, mensagem)
  If retorno Then
    bsShowMessage("Alteração não permitida." + mensagem, "E")
    CanContinue = False
  End If

  Set dll = Nothing

  Set qVerificaJaPago = Nothing
End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	If CommandID="BOTAOREVALIDAR" Then
		revalidar
	ElseIf CommandID="BOTAOREVISAR" Then
		Revisar
	ElseIf CommandID="BOTAOREVISARCOMANOTADM" Then
		RevisarComAnotacaoAdministrativa
	End If
End Sub

Public Sub Revisar
	Dim eventoSolicitBLL As CSBusinessComponent

	Set eventoSolicitBLL = BusinessComponent.CreateInstance("Benner.Saude.Atendimento.Business.SamAutorizEventoSolicitBLL, Benner.Saude.Atendimento.Business")

	eventoSolicitBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("AUTORIZACAO").AsInteger) 'autorizacao
	eventoSolicitBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger) 'eventoSolicitado

	eventoSolicitBLL.Execute("Revisar")

	Set eventoSolicitBLL = Nothing

	bsShowMessage("Revisão concluída com sucesso", "I")
End Sub

Public Sub RevisarComAnotacaoAdministrativa
	Dim eventoSolicitBLL As CSBusinessComponent

	Set eventoSolicitBLL = BusinessComponent.CreateInstance("Benner.Saude.Atendimento.Business.SamAutorizEventoSolicitBLL, Benner.Saude.Atendimento.Business")

	eventoSolicitBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("AUTORIZACAO").AsInteger) 'autorizacao
	eventoSolicitBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger) 'eventoSolicitado
	eventoSolicitBLL.AddParameter(pdtInteger, CurrentVirtualQuery.FieldByName("ANOTACAOADMINISTRATIVA").AsInteger) 'anotacaoAdministrativa
	eventoSolicitBLL.AddParameter(pdtAutomatic, CurrentVirtualQuery.FieldByName("ENVIARRELATORIORESPOSTA").AsString = "S") 'enviarNoRelatorioResposta
	eventoSolicitBLL.AddParameter(pdtString, CurrentVirtualQuery.FieldByName("OBSERVACAO").AsString) 'observacao

	eventoSolicitBLL.Execute("Revisar")

	Set eventoSolicitBLL = Nothing

	bsShowMessage("Revisão concluída com sucesso", "I")
End Sub
