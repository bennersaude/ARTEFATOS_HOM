'HASH: 0CE62385477D6B4B27A7EFBF8A1597FB
'#Uses "*bsShowMessage"
Dim gdCheque As Double

Public Sub CalcularValorJurosMulta()

	Dim vfValor As Double
	Dim vfJuro As Double
	Dim vfMulta As Double
	Dim vfCorrecao As Double
	Dim vfDesconto As Double
	Dim SFNBAIXA As Object
	Set SFNBAIXA = CreateBennerObject("SFNBAIXA.Documento")
	SFNBAIXA.BxCalcDocumento(CurrentSystem, _
							 CLng(SessionVar("HDOCUMENTO")), _
							 CurrentQuery.FieldByName("DATABAIXA").AsDateTime, _
							 CurrentQuery.FieldByName("TESOURARIA").AsInteger, _
							 vfValor, _
							 vfJuro, _
							 vfMulta, _
							 vfCorrecao, _
							 vfDesconto)
	Set SFNBAIXA = Nothing

	CurrentQuery.FieldByName("VALORCALC").AsFloat = vfValor
	CurrentQuery.FieldByName("JUROCALC").AsFloat = vfJuro
	CurrentQuery.FieldByName("MULTACALC").AsFloat = vfMulta
	CurrentQuery.FieldByName("CORRECAOCALC").AsFloat = vfCorrecao
	CurrentQuery.FieldByName("DESCONTOCALC").AsFloat = vfDesconto
	CurrentQuery.FieldByName("VALORINFORMADO").AsFloat = vfValor
	CurrentQuery.FieldByName("JUROINFORMADO").AsFloat = vfJuro
	CurrentQuery.FieldByName("MULTAINFORMADO").AsFloat = vfMulta
	CurrentQuery.FieldByName("CORRECAOINFORMADO").AsFloat = vfCorrecao
	CurrentQuery.FieldByName("DESCONTOINFORMADO").AsFloat = vfDesconto
	CurrentQuery.FieldByName("TOTALCALC").AsFloat = CurrentQuery.FieldByName("VALORCALC").AsFloat + _
													CurrentQuery.FieldByName("JUROCALC").AsFloat + _
													CurrentQuery.FieldByName("MULTACALC").AsFloat + _
													CurrentQuery.FieldByName("CORRECAOCALC").AsFloat - _
													CurrentQuery.FieldByName("DESCONTOCALC").AsFloat

	CalculaTotalInformado

End Sub

Public Sub CORRECAOINFORMADO_OnExit()
	CalculaTotalInformado
End Sub

Public Sub DATABAIXA_OnExit()
  CalcularValorJurosMulta
End Sub

Public Sub DESCONTOINFORMADO_OnExit()
	CalculaTotalInformado
End Sub

Public Sub JUROINFORMADO_OnExit()
	CalculaTotalInformado
End Sub

Public Sub MULTAINFORMADO_OnExit()
	CalculaTotalInformado
End Sub

Public Sub TABLE_AfterInsert()
	Dim qDocumento As BPesquisa
	Set qDocumento = NewQuery

	qDocumento.Clear
	qDocumento.Add("SELECT * FROM SFN_DOCUMENTO WHERE HANDLE = :HANDLE")
	qDocumento.ParamByName("HANDLE").AsInteger = CLng(SessionVar("HDOCUMENTO"))
	qDocumento.Active = True

	CurrentQuery.FieldByName("NOME").AsString = qDocumento.FieldByName("NOME").AsString
	CurrentQuery.FieldByName("NUMERO").AsInteger = qDocumento.FieldByName("NUMERO").AsInteger
	CurrentQuery.FieldByName("DATAVENCIMENTO").AsDateTime = qDocumento.FieldByName("DATAVENCIMENTO").AsDateTime
	CurrentQuery.FieldByName("DATAEMISSAO").AsDateTime = qDocumento.FieldByName("DATAEMISSAO").AsDateTime


	If ((qDocumento.FieldByName("VALOR").AsFloat - qDocumento.FieldByName("BAIXAVALOR").AsFloat > 0) Or _
		((qDocumento.FieldByName("VALOR").AsFloat = 0) And (qDocumento.FieldByName("BAIXAVALOR").AsFloat = 0))) Then

		CalcularValorJurosMulta

	End If

	qDocumento.Active = False

	Set qDocumento = Nothing
End Sub

Public Sub Baixar
	Dim SFNBAIXA As Object
	Dim vsMensagem As String
	Dim viRetorno As Long
	Set SFNBAIXA = CreateBennerObject("SFNBAIXA.Documento")

	StartTransaction

	viRetorno = SFNBAIXA.BaixarDocumento(CurrentSystem, _
										 CLng(SessionVar("HDOCUMENTO")), _
										 CurrentQuery.FieldByName("DATABAIXA").AsDateTime, _
										 CurrentQuery.FieldByName("DATACONTABIL").AsDateTime, _
										 CurrentQuery.FieldByName("TESOURARIA").AsInteger, _
										 CurrentQuery.FieldByName("VALORINFORMADO").AsFloat, _
										 CurrentQuery.FieldByName("JUROINFORMADO").AsFloat, _
										 CurrentQuery.FieldByName("MULTAINFORMADO").AsFloat, _
										 CurrentQuery.FieldByName("CORRECAOINFORMADO").AsFloat, _
										 CurrentQuery.FieldByName("DESCONTOINFORMADO").AsFloat, _
										 True, _
										 0, _
										 0, _
										 0, _
										 0, _
										 0, _
										 0, _
										 0, _
										 0, _
										 0, _
										 gdCheque, _
										 CurrentQuery.FieldByName("BANCO").AsInteger, _
										 CurrentQuery.FieldByName("AGENCIA").AsInteger, _
										 CurrentQuery.FieldByName("BAIXAMOTIVO").AsString, _
										 CurrentQuery.FieldByName("NOME").AsString, _
										 vsMensagem)

	Set SFNBAIXA = Nothing

	Select Case viRetorno
		Case -1
			bsShowMessage("Processo abortado pelo usuário!", "I")
			Rollback
		Case 0
			bsShowMessage("Baixa de documento efetuada com sucesso!", "I")
            Commit
		Case 1
          If vsMensagem <> "" Then
			bsShowMessage(vsMensagem, "I")
            Rollback
          End If
	End Select
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("TOTALINFORMADO").AsFloat < 0) Then
		bsShowMessage("O valor total não deve ser menor que zero!", "E")

		CanContinue = False

		Exit Sub
	End If

	If (CurrentQuery.FieldByName("DATABAIXA").AsDateTime > CurrentQuery.FieldByName("DATACONTABIL").AsDateTime) Then
		bsShowMessage("Data contábil não pode ser menor que data para baixa!", "E")

		CanContinue = False

		Exit Sub
	End If


	If (CDate(CurrentQuery.FieldByName("DATABAIXA").AsDateTime) < CDate(Format(CurrentQuery.FieldByName("DATAEMISSAO").AsDateTime, "DD/MM/YYYY")) ) Then
		bsShowMessage("Data não pode ser inferior a data de emissão!", "E")

		CanContinue = False

		Exit Sub
	End If

	If CurrentQuery.FieldByName("TESOURARIA").IsNull Then
		bsShowMessage("Tesouraria não informada", "E")

		CanContinue = False

		Exit Sub
	End If

	If (Not CurrentQuery.FieldByName("NUMEROCHEQUE").IsNull) Then
		gdCheque = CurrentQuery.FieldByName("NUMEROCHEQUE").AsInteger
	End If

        Baixar
End Sub
Public Sub CalculaTotalInformado
		CurrentQuery.FieldByName("TOTALINFORMADO").AsFloat = CurrentQuery.FieldByName("VALORINFORMADO").AsFloat + _
														 CurrentQuery.FieldByName("JUROINFORMADO").AsFloat + _
														 CurrentQuery.FieldByName("MULTAINFORMADO").AsFloat + _
														 CurrentQuery.FieldByName("CORRECAOINFORMADO").AsFloat - _
														 CurrentQuery.FieldByName("DESCONTOINFORMADO").AsFloat

End Sub

Public Sub TOTALINFORMADO_OnExit()
	CalculaTotalInformado
End Sub

Public Sub VALORINFORMADO_OnExit()
	CalculaTotalInformado
End Sub


Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Sql As Object

  Set Sql = NewQuery
  Sql.Add("SELECT COUNT (F.HANDLE) FATPROVISIAO                         ")
  Sql.Add("  FROM SFN_FATURA F											")
  Sql.Add("  JOIN SFN_DOCUMENTO_FATURA DF ON F.HANDLE = DF.FATURA		")
  Sql.Add("  JOIN SFN_DOCUMENTO D ON D.HANDLE = DOCUMENTO				")
  Sql.Add("  JOIN SIS_TIPOFATURAMENTO T ON T.HANDLE = F.TIPOFATURAMENTO ")
  Sql.Add(" WHERE T.CODIGO = 500										")
  Sql.Add("   AND D.HANDLE = " + SessionVar("HDOCUMENTO"))
  Sql.Active = True

  If Sql.FieldByName("FATPROVISIAO").AsInteger > 0 Then
    bsShowMessage("Operação não permitida para uma fatura de Provisão", "E")
    CanContinue = False
  End If

  Set Sql = Nothing
  Set qTipoFatura = Nothing
End Sub
