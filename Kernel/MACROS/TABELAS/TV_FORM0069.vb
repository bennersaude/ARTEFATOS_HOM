'HASH: 9743BCB41B30AAEDAAECE0C8DB1590D0
 
'#Uses "*bsShowMessage"

Dim	pHandleBanco As Long
Dim	pHandleAgencia As Long
Dim pDescAgencia As String
Dim	pDescBanco As String
Dim	pContaNumero As String
Dim	pContaDV As String
Dim	pCpfCnpj As String
Dim	pNome As String

Public Sub TABLE_AfterInsert()
	Dim qVencimento As BPesquisa

	Set qVencimento = NewQuery


	qVencimento.Clear
	qVencimento.Add("SELECT DATAVENCIMENTO, DATAEMISSAO FROM SFN_FATURA WHERE HANDLE = :HANDLE")
	qVencimento.ParamByName("HANDLE").AsInteger = CLng(SessionVar("HFATURA"))
	qVencimento.Active = True

	CurrentQuery.FieldByName("VENCIMENTOANTERIOR").Value = qVencimento.FieldByName("DATAVENCIMENTO").AsDateTime
	CurrentQuery.FieldByName("EMISSAOANTERIOR").Value = qVencimento.FieldByName("DATAEMISSAO").AsDateTime

	Dim SAMCONTAFINANC As Object
	Set SAMCONTAFINANC = CreateBennerObject("SamContaFinanceira.Consulta")

	SAMCONTAFINANC.ExibeContaFinanceira(CurrentSystem, _
										CLng(SessionVar("HFATURA")), _
										pHandleBanco, _
										pHandleAgencia, _
										pDescBanco, _
										pDescAgencia, _
										pContaNumero, _
										pContaDV, _
										pCpfCnpj, _
										pNome)



	CurrentQuery.FieldByName("BANCO").Value = pHandleBanco

	CurrentQuery.FieldByName("AGENCIA").Value = pHandleAgencia

	CurrentQuery.FieldByName("CONTACORRENTENUMERO").Value = pContaNumero

	CurrentQuery.FieldByName("CONTACORRENTEDV").AsString = pContaDV

	CurrentQuery.FieldByName("CONTACORRENTECPFCNPJ").AsString = pCpfCnpj

	CurrentQuery.FieldByName("CONTACORRENTENOME").AsString = pNome

	qVencimento.Active = False
	Set qVencimento = Nothing

End Sub

Public Sub TABLE_AfterPost()
	Dim vsMensagem As String


  Dim Sql As Object
  Set Sql = NewQuery
  Sql.Add("SELECT CODIGO                                                    ")
  Sql.Add("  FROM SIS_TIPOFATURAMENTO TF                                    ")
  Sql.Add("  JOIN SFN_FATURA           F ON (TF.HANDLE = F.TIPOFATURAMENTO) ")
  Sql.Add(" WHERE F.HANDLE = :HNDFAT                                        ")
  Sql.ParamByName("HNDFAT").AsInteger = CLng(SessionVar("HFATURA"))
  Sql.Active = True

  If Sql.FieldByName("CODIGO").AsInteger = 500 Then
	bsShowMessage("Operação não permitida para uma fatura de Provisão", "I")
	Exit Sub
  Else
    Dim SAMCONTAFINANC As Object
	Set SAMCONTAFINANC = CreateBennerObject("SamContaFinanceira.Consulta")

	vsMensagem = SAMCONTAFINANC.AlterarDataVencimentoFatura(CurrentSystem, _
															CLng(SessionVar("HFATURA")), _
															CurrentQuery.FieldByName("NOVOVENCIMENTO").AsDateTime, _
															CurrentQuery.FieldByName("NOVAEMISSAO").AsDateTime, _
															CurrentQuery.FieldByName("OBSERVACAO").AsString, _
															CurrentQuery.FieldByName("BANCO").Value, _
															CurrentQuery.FieldByName("AGENCIA").Value, _
															CurrentQuery.FieldByName("CONTACORRENTENUMERO").AsString, _
															CurrentQuery.FieldByName("CONTACORRENTEDV").AsString, _
															CurrentQuery.FieldByName("CONTACORRENTECPFCNPJ").AsString, _
															CurrentQuery.FieldByName("CONTACORRENTENOME").AsString)

	If (vsMensagem <> "") Then
		bsShowMessage(vsMensagem, "I")
	End If
  End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  CanContinue = True

  If pHandleBanco <> 0 Then
    If (CurrentQuery.FieldByName("BANCO").Value <= 0) Then
      CanContinue = False
      CancelDescription = "Favor preencher o campo 'Banco'."
      Exit Sub
    End If
	If (CurrentQuery.FieldByName("AGENCIA").Value <= 0) Then
      CanContinue = False
      CancelDescription = "Favor preencher o campo 'Agência'."
      Exit Sub
    End If
    If (Trim(CurrentQuery.FieldByName("CONTACORRENTENUMERO").AsString) = "") Then
      CanContinue = False
      CancelDescription = "Favor preencher o campo 'Número Conta Corrente'."
      Exit Sub
    End If
    If (Trim(CurrentQuery.FieldByName("CONTACORRENTEDV").AsString) = "") Then
      CanContinue = False
      CancelDescription = "Favor preencher o campo 'DV Conta Corrente'."
      Exit Sub
    End If
    If (Trim(CurrentQuery.FieldByName("CONTACORRENTECPFCNPJ").AsString) = "") Then
      CanContinue = False
      CancelDescription = "Favor preencher o campo 'CPF/CNPJ'."
      Exit Sub
    End If
    If (Trim(CurrentQuery.FieldByName("CONTACORRENTENOME").AsString) = "") Then
      CanContinue = False
      CancelDescription = "Favor preencher o campo 'Nome'."
      Exit Sub
    End If
  End If

End Sub
