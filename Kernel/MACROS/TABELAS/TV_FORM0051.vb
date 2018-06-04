'HASH: E71FC43F78C506DE3B01D15C171EE58A
'#Uses "*bsShowMessage"

Dim qFatura As BPesquisa

Public Sub TABLE_AfterInsert()

	Set qFatura = NewQuery

	qFatura.Clear
	qFatura.Add("SELECT F.*,                                                 ")
    qFatura.Add("       FL.TESOURARIALANC                                    ")
    qFatura.Add("  FROM SFN_FATURA F                                         ")
    qFatura.Add("  JOIN SFN_FATURA_LANC FL ON (F.HANDLE = FL.FATURA)         ")
    qFatura.Add(" WHERE F.HANDLE = :HANDLE                                   ")
	qFatura.ParamByName("HANDLE").AsString = SessionVar("HFATURA")
	qFatura.Active = True

	CurrentQuery.FieldByName("NUMERO").AsInteger = qFatura.FieldByName("NUMERO").AsInteger
	CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger = qFatura.FieldByName("TIPOFATURAMENTO").AsInteger
	CurrentQuery.FieldByName("DATAEMISSAO").AsDateTime = qFatura.FieldByName("DATAEMISSAO").AsDateTime
	CurrentQuery.FieldByName("DATAVENCIMENTO").AsDateTime = qFatura.FieldByName("DATAVENCIMENTO").AsDateTime
	CurrentQuery.FieldByName("DATACONTABILFATURA").AsDateTime = qFatura.FieldByName("DATACONTABIL").AsDateTime
	CurrentQuery.FieldByName("VALOR").AsFloat = qFatura.FieldByName("VALOR").AsFloat
	CurrentQuery.FieldByName("SALDO").AsFloat = qFatura.FieldByName("SALDO").AsFloat
	CurrentQuery.FieldByName("NATUREZA").AsString = qFatura.FieldByName("NATUREZA").AsString
	CurrentQuery.FieldByName("REGRAFINANCEIRA").AsInteger = qFatura.FieldByName("REGRAFINANCEIRA").AsInteger

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim Sql As Object

  Set Sql = NewQuery
  Sql.Add("SELECT CODIGO FROM SIS_TIPOFATURAMENTO WHERE HANDLE = :HANDLE")
  Sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("TIPOFATURAMENTO").AsInteger
  Sql.Active = True

  If Sql.FieldByName("CODIGO").AsInteger = 500 Then
	bsShowMessage("Operação não permitida para uma fatura de Provisão", "E")
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

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim vTesLancAntas As Long

  vTesLancAnt = qFatura.FieldByName("TESOURARIALANC").AsInteger
  qFatura.First
  While Not qFatura.EOF
    If vTesLancAnt <> qFatura.FieldByName("TESOURARIALANC").AsInteger Then
      BsShowMessage("Fatura com baixa parcial/total só pode ser estornada pelo Documento (quando existir) ou pela Tesouraria!", "E")
      CanContinue = False
	  Exit Sub
    End If

    vTesLancAnt = qFatura.FieldByName("TESOURARIALANC").AsInteger
    qFatura.Next
  Wend

    SessionVar("vForm") = "S"

	Dim SFNBAIXA As Object
	Dim viRetorno As Long
	Dim vsMensagem As String
	Set SFNBAIXA = CreateBennerObject("SFNBAIXA.Documento")



	viRetorno = SFNBAIXA.EstornoBaixaFat(CurrentSystem, _
										 CLng(SessionVar("HFATURA")), _
										 CurrentQuery.FieldByName("DATAESTORNO").AsDateTime, _
										 CurrentQuery.FieldByName("DATACONTABIL").AsDateTime, _
										 CurrentQuery.FieldByName("HISTORICO").AsString, _
										 CurrentQuery.FieldByName("TESOURARIA").AsInteger, _
										 qFatura.FieldByName("TESOURARIALANC").AsInteger, _
										 vsMensagem)

	Select Case viRetorno
		Case -1
			bsShowMessage("Processo abortado pelo usuário!", "E")
			CanContinue = False
			Exit Sub
		Case 0
			bsShowMessage("Estorno realizado com sucesso!", "I")
		Case 1
			bsShowMessage(vsMensagem, "E")
			CanContinue = False
			Exit Sub
	End Select

	Set SFNBAIXA = Nothing
End Sub
