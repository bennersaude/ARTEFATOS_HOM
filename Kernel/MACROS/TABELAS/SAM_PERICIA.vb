'HASH: 8D34837E9C9C35275BADFA263E362EC3
'macro: SAM_PERICIA

'#uses "*CriaTabelaTemporariaSqlServer"
'#uses "*bsShowMessage"

Option Explicit


Public Function verificaDigitacao As String
  Dim problemas As String
  Dim dias As Integer
  problemas = ""

  Dim sql As BPesquisa
  Set sql=NewQuery
  sql.Add("SELECT A.DATAAUTORIZACAO FROM SAM_AUTORIZ A WHERE A.HANDLE=:HANDLE")
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
  sql.Active=True

  dias = DateDiff("d", CurrentQuery.FieldByName("DATAPEDIDO").AsDateTime, ServerDate)

  If CurrentQuery.FieldByName("DATAPEDIDO").IsNull Then
    problemas = problemas + "Data de pedido da Pericia obrigatória." +Chr(13)
  Else
   If (CurrentQuery.FieldByName("DATAPEDIDO").AsDateTime < sql.FieldByName("DATAAUTORIZACAO").AsDateTime) Or (dias < 0) Then
      problemas = problemas + "Data de pedido deve ser entre a " + Chr(13) + "data da autorização (" + Format(sql.FieldByName("DATAAUTORIZACAO").AsString, "dd/mm/yyyy" ) + ") e data atual (" + Format(CStr(ServerDate), "dd/mm/yyyy" ) + ")" + Chr(13)
    End If
  End If

  If CurrentQuery.FieldByName("MOTIVOPERICIA").IsNull Then
    problemas = problemas + "Motivo de Pericia obrigatório." +Chr(13)
  End If

  If (CurrentQuery.FieldByName("TABREGRA").AsInteger = 2) And (CurrentQuery.FieldByName("MOTIVOPARECER").IsNull) Then
    problemas = problemas + "Motivo de parecer obrigatório." +Chr(13)
  End If

  If (CurrentQuery.FieldByName("PARECER").AsString = "30/12/1899") Or (CurrentQuery.FieldByName("PARECER").AsString = "00/00/0000") Then
    CurrentQuery.FieldByName("PARECER").Clear
  End If

  If (((CurrentQuery.FieldByName("PARECER").IsNull) And (Not CurrentQuery.FieldByName("PARECERDATA").IsNull)) Or _
  ((Not CurrentQuery.FieldByName("PARECER").IsNull) And (CurrentQuery.FieldByName("PARECERDATA").IsNull))) Then

    problemas = problemas + "Parecer deve ser informado com a data." +Chr(13)
  End If

  If ((CurrentQuery.FieldByName("PARECER").AsString <> "") And (Not CurrentQuery.FieldByName("PARECERDATA").IsNull) And _
    (CurrentQuery.FieldByName("AUDITOR").IsNull)) Then

    problemas = problemas + "Parecer exige que seja informado o perito." +Chr(13)
  End If

  If CurrentQuery.FieldByName("FILIALORIGEM").IsNull Then
    problemas = problemas + "Filial de origem obrigatória." +Chr(13)
  End If
  If CurrentQuery.FieldByName("FILIALDESTINO").IsNull Then
    problemas = problemas + "Filial de destino obrigatória." +Chr(13)
  End If


  verificaDigitacao = problemas

  Set sql = Nothing
End Function

Public Sub TABLE_AfterInsert()
  Dim qSQL As Object
  Set qSQL = NewQuery

  qSQL.Add("SELECT FILIALPADRAO")
  qSQL.Add("FROM Z_GRUPOUSUARIOS")
  qSQL.Add("WHERE HANDLE = :HUSUARIO")
  qSQL.ParamByName("HUSUARIO").AsInteger = CurrentUser
  qSQL.Active = True

  If Not qSQL.FieldByName("FILIALPADRAO").IsNull Then
    CurrentQuery.FieldByName("FILIALORIGEM").AsInteger = qSQL.FieldByName("FILIALPADRAO").AsInteger
  End If

  qSQL.Clear
  qSQL.Add("SELECT BEN.FILIALCUSTO")
  qSQL.Add("FROM SAM_AUTORIZ AUT")
  qSQL.Add("JOIN SAM_BENEFICIARIO BEN ON BEN.HANDLE = AUT.BENEFICIARIO")
  qSQL.Add("WHERE AUT.HANDLE = :HAUTORIZACAO")
  qSQL.ParamByName("HAUTORIZACAO").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
  qSQL.Active = True

  If Not qSQL.FieldByName("FILIALCUSTO").IsNull Then
    CurrentQuery.FieldByName("FILIALDESTINO").AsInteger = qSQL.FieldByName("FILIALCUSTO").AsInteger
  End If

  Set qSQL = Nothing
End Sub

Public Sub TABLE_AfterPost()
	WriteBDebugMessage("SAM_PERICIA.TABLE_AfterPost - Início")
	auditar
	reverteNegacaoPericia
	revalidarAutorizacao
	WriteBDebugMessage("SAM_PERICIA.TABLE_AfterPost - Fim")
End Sub

Public Sub TABLE_AfterScroll()
  If WebMode Then
	AUDITOR.WebLocalWhere = " A.SITUACAO = 'A' "
  Else
	AUDITOR.LocalWhere = " A.SITUACAO = 'A' "
  End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  If VisibleMode Then
    CanContinue = False
    bsShowMessage("Exclusão de perícia permitida apenas pela interface de Autorização", "E")
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  If VisibleMode Then
    CanContinue = False
    bsShowMessage("Alteração de perícia permitida apenas pela interface de Autorização", "E")
  End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If VisibleMode Then
    CanContinue = False
    bsShowMessage("Inclusão de perícia permitida apenas pela interface de Autorização", "E")
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	WriteBDebugMessage("SAM_PERICIA.TABLE_BeforePost - Início")
	CriaTabelaTemporariaSqlServer
	Dim problemas As String
	problemas = verificaDigitacao
	If problemas <> "" Then
		CancelDescription = problemas
		CanContinue=False
	End If

	If (CurrentQuery.State = 1) Then
		Dim sqlBuscaPericia As BPesquisa
		Set sqlBuscaPericia = NewQuery

		sqlBuscaPericia.Add("SELECT 1 FROM SAM_PERICIA WHERE AUTORIZACAO = :autorizacao ")
		sqlBuscaPericia.ParamByName("autorizacao").AsInteger = CurrentQuery.FieldByName("autorizacao").AsInteger
		sqlBuscaPericia.Active = True

		If (Not sqlBuscaPericia.EOF) Then
			WriteBDebugMessage("SAM_PERICIA.TABLE_BeforePost - Foi cadastrada uma perícia recentemente")
			bsshowmessage("Foi cadastrada uma perícia recentemente","E")
			CanContinue = False
		End If

		Set sqlBuscaPericia = Nothing
	End If
	WriteBDebugMessage("SAM_PERICIA.TABLE_BeforePost - Fim")
End Sub


Public Sub auditar
	WriteBDebugMessage("SAM_PERICIA.auditar - Início")
	Dim vLog As String
	Dim sql As BPesquisa
	Set sql=NewQuery
	sql.Add("SELECT * FROM SAM_PERICIA WHERE HANDLE= :HANDLE AND SERVICOPROPRIO = 'N'")
	sql.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
	sql.Active=True

	vLog = "Atualização dos campos: " + Chr(13)
	Dim j As Integer
	For j = 0 To sql.FieldCount - 1
		vLog = vLog + Chr(13) + sql.Fields(j).FieldName + ": " + sql.FieldByName(sql.Fields(j).FieldName).AsString
	Next j
	vLog = vLog + Chr(13) + "na tabela SAM_PERICIA"
	WriteAudit("A", HandleOfTable("SAM_PERICIA"), CurrentQuery.FieldByName("HANDLE").AsInteger, vLog)
	Set sql = Nothing
	WriteBDebugMessage("SAM_PERICIA.auditar - Fim")
End Sub

Public Sub reverteNegacaoPericia
	WriteBDebugMessage("SAM_PERICIA.reverteNegacaoPericia - Início")
	Dim sql As BPesquisa
	Set sql=NewQuery
	sql.Add("SELECT REVERSAOAUTNEGPERICIA FROM SAM_PARAMETROSATENDIMENTO")
	sql.Active=True
	If sql.FieldByName("REVERSAOAUTNEGPERICIA").AsString = "S" Then
		Dim SP As BStoredProc
		Set SP=NewStoredProc
		SP.Name="BSAUT_VERIFICANEGPERICIA"
		SP.AddParam("P_AUTORIZACAO", ptInput,ftInteger)
		SP.AddParam("P_PROTOCOLOTRANSACAO", ptInput,ftInteger)
		SP.AddParam("P_PERICIA", ptInput,ftInteger)
		SP.AddParam("P_USUARIO", ptInput,ftInteger)
		SP.AddParam("P_EVENTOGERADO", ptInput,ftInteger)

		SP.ParamByName("P_AUTORIZACAO").AsInteger = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
		SP.ParamByName("P_PROTOCOLOTRANSACAO").AsInteger = CurrentQuery.FieldByName("PROTOCOLO").AsInteger
		SP.ParamByName("P_PERICIA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		SP.ParamByName("P_USUARIO").AsInteger = CurrentUser
		SP.ParamByName("P_EVENTOGERADO").Clear
		SP.ExecProc
	End If
	Set sql=Nothing
	WriteBDebugMessage("SAM_PERICIA.reverteNegacaoPericia - Fim")
End Sub

Public Sub revalidarAutorizacao
	WriteBDebugMessage("SAM_PERICIA.revalidarAutorizacao - Início")
	Dim retorno As Integer
	Dim mensagem As String
	Dim alertas As String
	Dim vHandleAutorizacao As Long
	Dim vHandleProtocoloTransacaoAUtoriz As Long

    vHandleAutorizacao = CurrentQuery.FieldByName("AUTORIZACAO").AsInteger
    vHandleProtocoloTransacaoAUtoriz = 0

	If RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ") > 0 Then
	  Dim qBuscaAutorizacao As Object
	  Set qBuscaAutorizacao = NewQuery

	  qBuscaAutorizacao.Clear
	  qBuscaAutorizacao.Add("SELECT AUTORIZACAO FROM SAM_PROTOCOLOTRANSACAOAUTORIZ WHERE HANDLE = :HANDLE")
	  qBuscaAutorizacao.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ")
	  qBuscaAutorizacao.Active = True

      vHandleAutorizacao = qBuscaAutorizacao.FieldByName("AUTORIZACAO").AsInteger
      vHandleProtocoloTransacaoAUtoriz = RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ")

	  Set qBuscaAutorizacao = Nothing
	End If

	Dim dll As Object
	Set dll=CreateBennerObject("ca043.autorizacao")
	retorno = dll.revalidarAutorizacao(CurrentSystem, vHandleAutorizacao, vHandleProtocoloTransacaoAUtoriz, alertas, mensagem)
	Set dll=Nothing
	If retorno > 0 Then
		WriteBDebugMessage("SAM_PERICIA.revalidarAutorizacao - " + mensagem)
		InfoDescription = mensagem
	End If
	WriteBDebugMessage("SAM_PERICIA.revalidarAutorizacao - Fim")
End Sub
