'HASH: F288F6902711348549C2FB772FC5922B
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
	CurrentQuery.FieldByName("VALORCANCELADO").AsFloat = qFatura.FieldByName("VALOR").AsFloat
	CurrentQuery.FieldByName("OBSERVACAO").AsString = qFatura.FieldByName("OBSERVACAO").AsString

	qFatura.Active = False

	Set qFatura = Nothing
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  Dim SAMCONTAFINANCEIRA As Object
  Dim vsDocumentos As String
  Dim vsMensagemRetorno As String

  Dim qTipoFatura As Object

  Set qTipoFatura = NewQuery
  qTipoFatura.Add("SELECT TIPOFATURAMENTO FROM SFN_FATURA WHERE HANDLE = " + SessionVar("HFATURA"))
  qTipoFatura.Active = True


  Dim Sql As Object

  Set Sql = NewQuery
  Sql.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE = " + qTipoFatura.FieldByName("TIPOFATURAMENTO").AsString)
  Sql.Active = True

  If Sql.FieldByName("CODIGO").AsInteger = 500 Then
    bsShowMessage("Operação não permitida para uma fatura de Provisão", "E")
    CanContinue = False
  End If

  Set Sql = Nothing
  Set qTipoFatura = Nothing

  Set SAMCONTAFINANCEIRA = CreateBennerObject("SAMCONTAFINANCEIRA.Consulta")

  If Not SAMCONTAFINANCEIRA.VerificaDocumentos(CurrentSystem,CLng(SessionVar("HFATURA")),vsDocumentos,vsMensagemRetorno) Then
    BsShowMessage(vsMensagemRetorno,"E")
    CanContinue = False
  End If

  Set SAMCONTAFINANCEIRA = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim EspecificoDll As Object
	Dim vsDocumentos As String

	Set EspecificoDll = CreateBennerObject("ESPECIFICO.uESPECIFICO")

	vsDocumentos = EspecificoDll.FIN_VerificaDocumentoFatura(CurrentSystem, _
															 CLng(SessionVar("HFATURA")))

	Set EspecificoDll = Nothing

	If (Not vsDocumentos = "") Then
		bsShowMessage("Existe documento aberto para esta fatura. Não é possível continuar !", "E")

		CanContinue = False

		Exit Sub
	End If

	If CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime < CurrentQuery.FieldByName("DATAEMISSAOFATURA").AsDateTime Then
      bsShowMessage("A data de cancelamento deve ser superior à data de emissão da fatura.", "I")

      DATACANCELAMENTO.SetFocus

      CanContinue = False

      Exit Sub
  	End If


	Dim SFNCANCEL As Object
	Dim vsMensagem As String
	Dim viResult As Long
	Set SFNCANCEL = CreateBennerObject("SFNCANCEL.Cancelamento")

	On Error GoTo erro

	If Not InTransaction Then StartTransaction

	viResult = SFNCANCEL.CancelaFatura(CurrentSystem, _
									   CurrentQuery.FieldByName("VALORCANCELADO").AsFloat, _
									   CLng(SessionVar("HFATURA")), _
									   CurrentQuery.FieldByName("MOTIVOCANCELAMENTO").AsString, _
									   CurrentQuery.FieldByName("DATACANCELAMENTO").AsDateTime, _
									   CurrentQuery.FieldByName("DATACONTABIL").AsDateTime, _
									   120, _
									   vsMensagem)

	Select Case viResult
		Case -1
			If vsMensagem <> "" Then
				bsShowMessage(vsMensagem, "I")
			Else
				bsShowMessage("Processo abortado pelo usuário.", "I")
			End If
		Case 0
			bsShowMessage("Fatura cancelada com sucesso.", "I")
		Case 1
			bsShowMessage(vsMensagem, "I")
	End Select

	If InTransaction Then Commit

	Exit Sub
	erro:
		bsShowMessage(Err.Description, "I")
		If InTransaction Then
			Rollback
		End If

End Sub
