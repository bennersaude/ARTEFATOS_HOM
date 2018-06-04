'HASH: 77E66D60BF9B2C990BAC19F7D227CAAD
'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterDelete()
	revalidarAutorizacao
End Sub

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
  qSQL.Add("  FROM SAM_PROTOCOLOTRANSACAOAUTORIZ PROTO")
  qSQL.Add("  JOIN SAM_AUTORIZ                   AUT   ON (AUT.HANDLE = PROTO.AUTORIZACAO)")
  qSQL.Add("  JOIN SAM_BENEFICIARIO BEN ON BEN.HANDLE = AUT.BENEFICIARIO")
  qSQL.Add(" WHERE PROTO.HANDLE = :HPROTOCOLO")
  qSQL.ParamByName("HPROTOCOLO").AsInteger = CLng(RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ"))
  qSQL.Active = True

  If Not qSQL.FieldByName("FILIALCUSTO").IsNull Then
    CurrentQuery.FieldByName("FILIALDESTINO").AsInteger = qSQL.FieldByName("FILIALCUSTO").AsInteger
  End If

  Set qSQL = Nothing

End Sub

Public Sub TABLE_AfterPost()
	revalidarAutorizacao
End Sub

Public Sub TABLE_AfterScroll()

  If RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ") > 0 Then

    If WebMode Then
      EVENTOGERADO.WebLocalWhere = "A.HANDLE IN (SELECT EG.HANDLE " + _
                                   "               FROM SAM_PERICIA PER " + _
                                   "               JOIN SAM_AUTORIZ_EVENTOGERADO EG ON EG.AUTORIZACAO = PER.AUTORIZACAO " + _
                                   "              WHERE EG.PROTOCOLOTRANSACAO =  " + CStr(RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ")) + _
                                   "                AND NOT EXISTS (SELECT X.EVENTOGERADO FROM SAM_PERICIA_EVENTO X WHERE X.EVENTOGERADO = EG.HANDLE))"
    Else
      EVENTOGERADO.LocalWhere    = "  HANDLE IN (SELECT EG.HANDLE " + _
                                   "               FROM SAM_PERICIA PER " + _
                                   "               JOIN SAM_AUTORIZ_EVENTOGERADO EG ON EG.AUTORIZACAO = PER.AUTORIZACAO " + _
                                   "              WHERE PER.HANDLE =  " + CStr(RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ")) + _
                                   "                AND NOT EXISTS (SELECT X.EVENTOGERADO FROM SAM_PERICIA_EVENTO X WHERE X.EVENTOGERADO = EG.HANDLE))"
    End If

  Else
    If WebMode Then
      EVENTOGERADO.WebLocalWhere = "A.HANDLE IN (SELECT EG.HANDLE " + _
                                   "               FROM SAM_PERICIA PER " + _
                                   "               JOIN SAM_AUTORIZ_EVENTOGERADO EG ON EG.AUTORIZACAO = PER.AUTORIZACAO " + _
                                   "              WHERE PER.HANDLE = @CAMPO(PERICIA) " + _
                                   "                AND NOT EXISTS (SELECT X.EVENTOGERADO FROM SAM_PERICIA_EVENTO X WHERE X.EVENTOGERADO = EG.HANDLE))"
    Else
      EVENTOGERADO.LocalWhere    = "  HANDLE IN (SELECT EG.HANDLE " + _
                                   "               FROM SAM_PERICIA PER " + _
                                   "               JOIN SAM_AUTORIZ_EVENTOGERADO EG ON EG.AUTORIZACAO = PER.AUTORIZACAO " + _
                                   "              WHERE PER.HANDLE = @PERICIA " + _
                                   "                AND NOT EXISTS (SELECT X.EVENTOGERADO FROM SAM_PERICIA_EVENTO X WHERE X.EVENTOGERADO = EG.HANDLE))"
    End If

  End If

  If WebMode Then
	AUDITOR.WebLocalWhere = " A.SITUACAO = 'A' "
  Else
	AUDITOR.LocalWhere = " SAM_AUDITOR.SITUACAO = 'A' "
  End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  If VisibleMode Then
    CanContinue = False
    bsShowMessage("Inclusão de eventos permitida apenas pela interface de Autorização", "E")
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim problemas As String
	problemas = validarDigitacao
	If problemas<>"" Then
		CancelDescription = problemas
  		CanContinue=False
	End If
End Sub

Public Function validarDigitacao As String
  Dim problemas As String
  problemas = ""

  Dim sql As BPesquisa
  Set sql=NewQuery
  sql.Add("SELECT DATAHORAINCLUSAO FROM SAM_PROTOCOLOTRANSACAOAUTORIZ A WHERE HANDLE=:HANDLE")
  sql.ParamByName("HANDLE").AsInteger = CLng(RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ"))
  sql.Active=True

  If CurrentQuery.FieldByName("MOTIVOPERICIA").IsNull Then
    problemas = problemas + "Motivo de Pericia obrigatório." +Chr(13)
  End If

  If (CurrentQuery.FieldByName("ACAOPERICIA").AsString = "N") And (CurrentQuery.FieldByName("MOTIVOPARECERAUDITORIA").IsNull) Then
    problemas = problemas + "Motivo de parecer obrigatório." +Chr(13)
  End If

  If (CurrentQuery.FieldByName("PARECERDATA").AsString = "30/12/1899") Or (CurrentQuery.FieldByName("PARECERDATA").AsString = "00/00/0000") Then
    CurrentQuery.FieldByName("PARECERDATA").Clear
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


  validarDigitacao = problemas

  Set sql = Nothing

End Function


Public Sub revalidarAutorizacao

	Dim retorno As Integer
	Dim mensagem As String
	Dim alertas As String
	Dim vHandleAutorizacao As Long
	Dim vHandleProtocoloTransacaoAUtoriz As Long

    vHandleAutorizacao = RecordHandleOfTable("SAM_AUTORIZ")
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
		InfoDescription = mensagem
	End If

End Sub
