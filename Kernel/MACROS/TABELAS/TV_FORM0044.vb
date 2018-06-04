'HASH: 93D8E97CDF1FE7670999EFA9F2A719FE
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	Dim qFatura As BPesquisa
	Set qFatura = NewQuery

	qFatura.Clear
	qFatura.Add("SELECT NUMERO,			")
	qFatura.Add("		TIPOFATURAMENTO,")
	qFatura.Add("		DATAEMISSAO,	")
	qFatura.Add("		DATACONTABIL,	")
	qFatura.Add("		DATAVENCIMENTO,	")
	qFatura.Add("		VALOR,			")
	qFatura.Add("		SALDO,			")
	qFatura.Add("		NATUREZA,		")
	qFatura.Add("		SITUACAO,		")
	qFatura.Add("		REGRAFINANCEIRA,")
	qFatura.Add("		OBSERVACAO		")
	qFatura.Add("  FROM SFN_FATURA		")
	qFatura.Add(" WHERE HANDLE = :HANDLE")
	qFatura.ParamByName("HANDLE").AsString = SessionVar("HFATURA")
	qFatura.Active = True

	CurrentQuery.FieldByName("NUMEROFATURA").AsInteger = qFatura.FieldByName("NUMERO").AsInteger
	CurrentQuery.FieldByName("TIPOFATURAMENTOFATURA").AsInteger = qFatura.FieldByName("TIPOFATURAMENTO").AsInteger
	CurrentQuery.FieldByName("DATAEMISSAOFATURA").AsDateTime = qFatura.FieldByName("DATAEMISSAO").AsDateTime
	CurrentQuery.FieldByName("DATACONTABILFATURA").AsDateTime = qFatura.FieldByName("DATACONTABIL").AsDateTime
	CurrentQuery.FieldByName("DATAVENCIMENTOFATURA").AsDateTime = qFatura.FieldByName("DATAVENCIMENTO").AsDateTime
	CurrentQuery.FieldByName("VALORFATURA").AsFloat = qFatura.FieldByName("VALOR").AsFloat
	CurrentQuery.FieldByName("SALDOFATURA").AsFloat = qFatura.FieldByName("SALDO").AsFloat
	CurrentQuery.FieldByName("NATUREZAFATURA").AsString = qFatura.FieldByName("NATUREZA").AsString
	CurrentQuery.FieldByName("SITUACAOFATURA").AsString = qFatura.FieldByName("SITUACAO").AsString
	CurrentQuery.FieldByName("REGRAFINANCEIRAFATURA").AsInteger = qFatura.FieldByName("REGRAFINANCEIRA").AsInteger
	CurrentQuery.FieldByName("OBSERVACAO").AsString = qFatura.FieldByName("OBSERVACAO").AsString

	qFatura.Active = False

	Set qFatura = Nothing
End Sub

Public Sub TABLE_AfterPost()
	Dim SFNCANCEL As Object
	Dim viResult As Integer
	Dim vsMensagem As String
	Set SFNCANCEL = CreateBennerObject("SFNCANCEL.Cancelamento")

	viResult = SFNCANCEL.EstornoCancFatura(CurrentSystem, _
										   CLng(SessionVar("HFATURA")), _
										   CurrentQuery.FieldByName("DATAESTORNO").AsDateTime, _
										   CurrentQuery.FieldByName("DATACONTABIL").AsDateTime, _
										   CurrentQuery.FieldByName("HISTORICO").AsString, _
										   vsMensagem)
	Select Case viResult
		Case -1
			bsShowMessage("Processo abortado pelo usuário.", "I")
		Case 0
			bsShowMessage("Estorno de cancelamento realizado com sucesso.", "I")
		Case 1
			bsShowMessage(vsMensagem, "I")
	End Select
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Sql As Object
  Set Sql = NewQuery

  If CurrentQuery.FieldByName("TIPOFATURAMENTOFATURA").AsString <> "" Then
    Sql.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE = " + CurrentQuery.FieldByName("TIPOFATURAMENTOFATURA").AsString)
    Sql.Active = True
    If Sql.FieldByName("CODIGO").AsInteger = 500 Then
      bsShowMessage("Operação não permitida para uma fatura de Provisão", "E")
      CanContinue = False
    End If
  End If

  Sql.Active = False
  Sql.Clear
  Sql.Add("SELECT COUNT(1) LANCANTEC                           ")
  Sql.Add("  FROM SFN_FATURA_LANC L                            ")
  Sql.Add("  JOIN SIS_OPERACAO    O ON (O.HANDLE = L.OPERACAO) ")
  Sql.Add(" WHERE O.CODIGO IN ('111', '112')                   ")
  Sql.Add("   AND L.FATURA = :HFATURA                          ")

  Sql.ParamByName("HFATURA").AsString = SessionVar("HFATURA")
  Sql.Active = True

  If Sql.FieldByName("LANCANTEC").AsInteger > 0 Then
    bsShowMessage("Operação não permitida para uma fatura com lançamentos de pré-pagamento (111, 112)", "E")
    CanContinue = False
  End If

  Set Sql = Nothing

  Dim SAMCONTAFINANCEIRA As Object
  Dim vsDocumentos As String
  Dim vsMensagemRetorno As String

  Set SAMCONTAFINANCEIRA = CreateBennerObject("SAMCONTAFINANCEIRA.Consulta")

  If Not SAMCONTAFINANCEIRA.VerificaDocumentos(CurrentSystem,CLng(SessionVar("HFATURA")),vsDocumentos,vsMensagemRetorno) Then
    BsShowMessage(vsMensagemRetorno,"E")
    CanContinue = False
  End If

  Set SAMCONTAFINANCEIRA = Nothing

End Sub
