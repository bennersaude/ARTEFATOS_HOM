'HASH: 5AFD04479E7C381922D2A79F7D9092B6
'MACRO: SAM_AUTORIZ_EVENTOGERADO
'#uses "*CriaTabelaTemporariaSqlServer"
'#uses "*bsShowMessage"
Option Explicit

Dim gsDadosEventoGeradoAntes As String
Dim gsDadosEventoGeradoDepois As String

Public Sub TABLE_AfterEdit()
  Dim DLL As Object
  Set DLL = CreateBennerObject("CA043.Autorizacao")
  gsDadosEventoGeradoAntes = DLL.ObterDadosEventoGerado(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set DLL = Nothing

  If CurrentQuery.FieldByName("PRAZOCLIENTE").IsNull Then
    Dim sp As BStoredProc

    Set sp = NewStoredProc

    sp.Name = "BS_74410031"
    sp.AddParam("p_Autorizacao",ptInput, ftInteger)
    sp.AddParam("p_Evento",ptInput, ftInteger)
    sp.AddParam("p_DataBase", ptInput,ftDateTime)
    sp.AddParam("p_PrazoCliente", ptOutput, ftDateTime)

    sp.ParamByName("p_Autorizacao").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
    sp.ParamByName("p_Evento").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    sp.ParamByName("p_DataBase").AsDateTime = CurrentQuery.FieldByName("DATAATENDIMENTO").AsDateTime
    sp.ExecProc

    CurrentQuery.FieldByName("PRAZOCLIENTE").AsDateTime = sp.ParamByName("p_PrazoCliente").AsDateTime

    Set sp = Nothing
  End If
End Sub

Public Sub TABLE_AfterPost()
  Dim viRetorno As Integer
  Dim vsMensagemErro As String

  Dim DLL As Object
  Set DLL = CreateBennerObject("CA043.Autorizacao")
  gsDadosEventoGeradoDepois = DLL.ObterDadosEventoGerado(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)
  viRetorno = DLL.EventoGeradoConsistenciasAposAlterar(CurrentSystem, gsDadosEventoGeradoAntes, gsDadosEventoGeradoDepois, vsMensagemErro)
  Set DLL = Nothing
  If viRetorno > 0 Then
    bsShowMessage(vsMensagemErro, "E")
  End If

  revalidar
End Sub

Public Sub TABLE_AfterScroll()

  RecordReadOnly = True

  BOTAOALERTA.Enabled = False
  BOTAOREVERTER.Enabled = False
  BOTAOREVALIDAR.Enabled = False
  BOTAOREVERTEROUTROUSUARIO.Enabled = False

  SessionVar("HANDLEEVENTOSOLICIT") = CurrentQuery.FieldByName("EVENTOSOLICITADO").AsString

  ROTULO2.Text = pegaRotulo2
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  RecordReadOnly = True
  If VisibleMode Then
    bsShowMessage("Operação permitida apenas pela Interface de Autorização!", "E")
    CanContinue = False
  Else
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
  If CurrentQuery.FieldByName("QTDAUTORIZADA").AsInteger < CurrentQuery.FieldByName("QTDPAGA").AsInteger Then
    bsShowMessage("Alteração não permitida. Quantidade solicitada menor que quantidade já paga para o evento", "E")
    CanContinue = False
  End If
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


Public Sub revalidar
	CriaTabelaTemporariaSqlServer
	Dim retorno As Integer
	Dim mensagem As String
	Dim alertas As String

	Dim DLL As Object
	Set DLL=CreateBennerObject("ca043.autorizacao")
	retorno = DLL.revalidarGerado(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, alertas, mensagem)
	Set DLL=Nothing
	If retorno > 0 Then
		InfoDescription = mensagem
	Else
		SessionVar("alertas") = alertas
		InfoDescription = "Revalidação concluída com sucesso"
	End If
End Sub


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

Public Sub Revisar
	Dim eventoGeradoBLL As CSBusinessComponent

	Set eventoGeradoBLL = BusinessComponent.CreateInstance("Benner.Saude.Atendimento.Business.SamAutorizEventoGeradoBLL, Benner.Saude.Atendimento.Business")

	eventoGeradoBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("AUTORIZACAO").AsInteger) 'autorizacao
	eventoGeradoBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("PROTOCOLOTRANSACAO").AsInteger) 'protocoloTransacaoAutorizacao
	eventoGeradoBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger) 'eventoGerado

	eventoGeradoBLL.Execute("Revisar")

	Set eventoGeradoBLL = Nothing

	bsShowMessage("Revisão concluída com sucesso", "I")
End Sub

Public Sub RevisarComAnotacaoAdministrativa
	Dim eventoGeradoBLL As CSBusinessComponent

	Set eventoGeradoBLL = BusinessComponent.CreateInstance("Benner.Saude.Atendimento.Business.SamAutorizEventoGeradoBLL, Benner.Saude.Atendimento.Business")

	eventoGeradoBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("AUTORIZACAO").AsInteger) 'autorizacao
	eventoGeradoBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("PROTOCOLOTRANSACAO").AsInteger) 'protocoloTransacaoAutorizacao
	eventoGeradoBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("EVENTOSOLICITADO").AsInteger) 'eventoSolicitado
	eventoGeradoBLL.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger) 'eventoGerado
	eventoGeradoBLL.AddParameter(pdtInteger, CurrentVirtualQuery.FieldByName("ANOTACAOADMINISTRATIVA").AsInteger) 'anotacaoAdministrativa
	eventoGeradoBLL.AddParameter(pdtAutomatic, CurrentVirtualQuery.FieldByName("ENVIARRELATORIORESPOSTA").AsString = "S") 'enviarNoRelatorioResposta
	eventoGeradoBLL.AddParameter(pdtString, CurrentVirtualQuery.FieldByName("OBSERVACAO").AsString) 'observacao

	eventoGeradoBLL.Execute("Revisar")

	Set eventoGeradoBLL = Nothing

	bsShowMessage("Revisão concluída com sucesso", "I")
End Sub
