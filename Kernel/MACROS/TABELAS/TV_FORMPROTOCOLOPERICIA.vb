'HASH: 4DC23895CAF28B5C72C1F4C8CD38F0CD
 
Option Explicit
'#Uses "*bsShowMessage"

Public Function retornaQueryPericiaEventoPendente As String

  retornaQueryPericiaEventoPendente = "SELECT PEVENTO.HANDLE                                                            " + _
                                      "  FROM SAM_PERICIA_EVENTO  PEVENTO                                               " + _
                                      " WHERE PEVENTO.EVENTOGERADO IN (SELECT HANDLE                                    " + _
                                      "                                  FROM SAM_AUTORIZ_EVENTOGERADO                  " + _
                                      "                                 WHERE PROTOCOLOTRANSACAO = :PROTOCOLOTRANSACAO) " + _
                                      "   AND PEVENTO.ACAOPERICIA = 'P' "

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

  Dim qAjustaPericiaEvento As Object
  Set qAjustaPericiaEvento = NewQuery

  qAjustaPericiaEvento.Clear
  qAjustaPericiaEvento.Add("UPDATE SAM_PERICIA_EVENTO ")
  qAjustaPericiaEvento.Add("   SET ")

  If CurrentQuery.FieldByName("TABREGRA").AsInteger = 1 Then ' Negar todos os eventos
    qAjustaPericiaEvento.Add("     ACAOPERICIA   = 'N' ")
  Else

    qAjustaPericiaEvento.Add("     ACAOPERICIA   = 'L' ") ' Liberar todos os eventos

    If CurrentQuery.FieldByName("MOTIVOPARECER").AsInteger > 0 Then ' Com motivo de parecer para liberar
      qAjustaPericiaEvento.Add("     , MOTIVOPARECERAUDITORIA = " + CurrentQuery.FieldByName("MOTIVOPARECER").AsString )
    End If

  End If

  If CurrentQuery.FieldByName("MOTIVOPERICIA").AsInteger > 0 Then ' Com motivo de perícia
    qAjustaPericiaEvento.Add("     , MOTIVOPERICIA = :MOTIVOPERICIA")
    qAjustaPericiaEvento.ParamByName("MOTIVOPERICIA").AsInteger =  CurrentQuery.FieldByName("MOTIVOPERICIA").AsInteger
  End If

  If CurrentQuery.FieldByName("FILIALORIGEM").AsInteger > 0 Then ' Com filial origem
    qAjustaPericiaEvento.Add("     , FILIALORIGEM = :FILIALORIGEM ")
    qAjustaPericiaEvento.ParamByName("FILIALORIGEM").AsInteger =  CurrentQuery.FieldByName("FILIALORIGEM").AsInteger
  End If

  If CurrentQuery.FieldByName("FILIALDESTINO").AsInteger > 0 Then ' Com filial destino
    qAjustaPericiaEvento.Add("     , FILIALDESTINO = :FILIALDESTINO ")
    qAjustaPericiaEvento.ParamByName("FILIALDESTINO").AsInteger =  CurrentQuery.FieldByName("FILIALDESTINO").AsInteger
  End If

  If CurrentQuery.FieldByName("AUDITOR").AsInteger > 0 Then ' Com Auditor
    qAjustaPericiaEvento.Add("     , AUDITOR = :AUDITOR ")
    qAjustaPericiaEvento.ParamByName("AUDITOR").AsInteger =  CurrentQuery.FieldByName("AUDITOR").AsInteger
  End If

  If CurrentQuery.FieldByName("PARECER").AsString <> "" Then ' Parecer
    qAjustaPericiaEvento.Add("     , PARECER = :PARECER ")
    qAjustaPericiaEvento.ParamByName("PARECER").AsString =  CurrentQuery.FieldByName("PARECER").AsString
  End If

  If Not CurrentQuery.FieldByName("PARECERDATA").IsNull Then ' Parecer data
    qAjustaPericiaEvento.Add("     , PARECERDATA = :PARECERDATA")
    qAjustaPericiaEvento.ParamByName("PARECERDATA").AsDateTime =  CurrentQuery.FieldByName("PARECERDATA").AsDateTime
  End If

  If CurrentQuery.FieldByName("OBSERVACOES").AsString <> "" Then
    qAjustaPericiaEvento.Add("     , OBSERVACOES = :OBSERVACOES ")
    qAjustaPericiaEvento.ParamByName("OBSERVACOES").AsString =  CurrentQuery.FieldByName("OBSERVACOES").AsString
  End If

  qAjustaPericiaEvento.Add(" WHERE HANDLE IN (" + retornaQueryPericiaEventoPendente + ")")
  qAjustaPericiaEvento.ParamByName("PROTOCOLOTRANSACAO").AsInteger = CLng(RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ"))
  qAjustaPericiaEvento.ExecSQL

  Set qAjustaPericiaEvento = Nothing


  revalidarAutorizacao

End Sub

Public Sub TABLE_AfterScroll()
  CurrentQuery.FieldByName("PROTOCOLOTRANSACAO").AsInteger = CLng(RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ"))
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  Dim qBuscaPericiaEventoPendente As Object
  Set qBuscaPericiaEventoPendente = NewQuery

  qBuscaPericiaEventoPendente.Clear
  qBuscaPericiaEventoPendente.Add(retornaQueryPericiaEventoPendente)
  qBuscaPericiaEventoPendente.ParamByName("PROTOCOLOTRANSACAO").AsInteger = CLng(RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ"))
  qBuscaPericiaEventoPendente.Active = True

  If qBuscaPericiaEventoPendente.FieldByName("HANDLE").AsInteger = 0 Then
    bsShowMessage("Nâo existem eventos pendentes de perícia para este protocolo!.", "E")
    CanContinue = False
    Exit Sub
  End If
  Set qBuscaPericiaEventoPendente = Nothing

End Sub


Public Function validarDigitacao As String
  Dim problemas As String
  problemas = ""

  Dim sql As BPesquisa
  Set sql=NewQuery
  sql.Add("SELECT DATAHORAINCLUSAO FROM SAM_PROTOCOLOTRANSACAOAUTORIZ A WHERE HANDLE=:HANDLE")
  sql.ParamByName("HANDLE").AsInteger = CLng(RecordHandleOfTable("SAM_PROTOCOLOTRANSACAOAUTORIZ"))
  sql.Active=True


  If CurrentQuery.FieldByName("DATAPEDIDO").IsNull Then
    problemas = problemas + "Data de pedido da Pericia obrigatória." +Chr(13)
  Else
   If (CurrentQuery.FieldByName("DATAPEDIDO").AsDateTime < sql.FieldByName("DATAHORAINCLUSAO").AsDateTime) Or _
      (CurrentQuery.FieldByName("DATAPEDIDO").AsDateTime > ServerNow) Then
      problemas = problemas + "Data de pedido deve ser entre a " + Chr(13) + "data da autorização (" + sql.FieldByName("DATAHORAINCLUSAO").AsString  + ") e data atual (" + CStr(ServerNow) + ")" + Chr(13)
    End If
  End If

  If CurrentQuery.FieldByName("MOTIVOPERICIA").IsNull Then
    problemas = problemas + "Motivo de Pericia obrigatório." +Chr(13)
  End If

  If (CurrentQuery.FieldByName("TABREGRA").AsInteger = 2) And (CurrentQuery.FieldByName("MOTIVOPARECER").IsNull) Then
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

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim problemas As String
  problemas = validarDigitacao
  If problemas<>"" Then
    bsShowMessage(problemas, "E")
    CanContinue=False
  End If
End Sub
