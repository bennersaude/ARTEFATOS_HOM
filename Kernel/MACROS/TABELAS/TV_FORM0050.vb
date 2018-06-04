'HASH: BE4986A80BBF45A7E7D9B8D725AE980F
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	Dim qDocumento As BPesquisa
	Set qDocumento = NewQuery

	qDocumento.Clear
	qDocumento.Add("SELECT * FROM SFN_DOCUMENTO WHERE HANDLE = :HANDLE")
	qDocumento.ParamByName("HANDLE").AsInteger = CLng(SessionVar("HDOCUMENTO"))
	qDocumento.Active = True

	CurrentQuery.FieldByName("NUMERO").AsInteger = qDocumento.FieldByName("NUMERO").AsInteger
	CurrentQuery.FieldByName("TIPODOCUMENTO").AsInteger = qDocumento.FieldByName("TIPODOCUMENTO").AsInteger
	CurrentQuery.FieldByName("DATAESTORNO").AsDateTime = ServerDate
	CurrentQuery.FieldByName("DATACONTABIL").AsDateTime = ServerDate
	CurrentQuery.FieldByName("DATAEMISSAO").AsDateTime = qDocumento.FieldByName("DATAEMISSAO").AsDateTime
	CurrentQuery.FieldByName("DATAVENCIMENTO").AsDateTime = qDocumento.FieldByName("DATAVENCIMENTO").AsDateTime
	If Not qDocumento.FieldByName("DATAVALIDADE").IsNull Then
	  CurrentQuery.FieldByName("DATAVALIDADE").AsDateTime = qDocumento.FieldByName("DATAVALIDADE").AsDateTime
	Else
	  CurrentQuery.FieldByName("DATAVALIDADE").Clear
	End If
	CurrentQuery.FieldByName("VALOR").AsFloat = qDocumento.FieldByName("VALOR").AsFloat
	CurrentQuery.FieldByName("VALORTOTAL").AsFloat = qDocumento.FieldByName("VALORTOTAL").AsFloat
	CurrentQuery.FieldByName("NATUREZA").AsString = qDocumento.FieldByName("NATUREZA").AsString
	CurrentQuery.FieldByName("REGRAFINANCEIRA").AsInteger = qDocumento.FieldByName("REGRAFINANCEIRA").AsInteger
End Sub

Public Sub TABLE_AfterPost()
	SessionVar("vForm") = "S"

	Dim SFNBAIXA As Object
	Dim viRetorno As Long
	Dim vsMensagem As String
	Set SFNBAIXA = CreateBennerObject("SFNBAIXA.Documento")

	viRetorno = SFNBAIXA.EstornoBaixaDoc(CurrentSystem, _
										 CLng(SessionVar("HDOCUMENTO")), _
										 CurrentQuery.FieldByName("DATAESTORNO").AsDateTime, _
										 CurrentQuery.FieldByName("DATACONTABIL").AsDateTime, _
										 CurrentQuery.FieldByName("HISTORICO").AsString, _
										 CurrentQuery.FieldByName("TESOURARIA").AsInteger, _
										 0, _
										 vsMensagem)

	Select Case viRetorno
		Case -1
		  If vsMensagem <> "" Then
		    Err.Raise(vbsUserException, "", vsMensagem + Chr(13) + "Operação cancelada!")
		  Else
			Err.Raise(vbsUserException, "", "Processo abortado pelo usuário!")
		  End If
		Case 0
			bsShowMessage("Estorno realizado com sucesso!", "I")
		Case 1
			Err.Raise(vbsUserException, "", vsMensagem + Chr(13) + "Operação cancelada!")
	End Select
End Sub


Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim qTipoFatura As Object

  Set qTipoFatura = NewQuery
  qTipoFatura.Add("SELECT F.HANDLE                                                   ")
  qTipoFatura.Add("  FROM SFN_DOCUMENTO D                                            ")
  qTipoFatura.Add("  JOIN SFN_DOCUMENTO_FATURA DF ON (D.HANDLE = DF.DOCUMENTO)       ")
  qTipoFatura.Add("  JOIN SFN_FATURA F ON (F.HANDLE = DF.FATURA)                     ")
  qTipoFatura.Add("  JOIN SIS_TIPOFATURAMENTO TF ON (TF.HANDLE = F.TIPOFATURAMENTO)  ")
  qTipoFatura.Add(" WHERE D.HANDLE = :DOC                                            ")
  qTipoFatura.Add("   AND TF.CODIGO = 500                                            ")

  qTipoFatura.ParamByName("DOC").AsString = SessionVar("HDOCUMENTO")
  qTipoFatura.Active = True

  If Not qTipoFatura.EOF Then
    bsShowMessage("Operação não permitida para uma fatura de Provisão", "E")
    CanContinue = False
  End If

  Set qTipoFatura = Nothing
End Sub
