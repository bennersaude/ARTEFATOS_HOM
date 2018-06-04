'HASH: 2B1AEB3F48FAFF4F37F1BCE6438AE1B7

'#Uses "*bsShowMessage"
'#USES "*formataCampoFiltroComVirgula"

Dim vSLocalWhere      As String
Dim viHContaFinOrigem As Long

Public Sub TABLE_AfterInsert()
  CurrentQuery.FieldByName("BENEFICIARIO").AsInteger = 0
End Sub

Public Sub TABLE_AfterScroll()
  If vSLocalWhere <> "" Then
    If WebMode Then
      FATURAS.WebLocalWhere = vSLocalWhere
    Else
      FATURAS.Where = vSLocalWhere
    End If
  End If

End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
  Dim vSLocalWhereBeneficiario As String
  Dim vSLocalWherePessoa As String

  Dim qBuscaContaFin As Object
  Set qBuscaContaFin = NewQuery
  qBuscaContaFin.Clear
  vSLocalWhere = ""
  vSLocalWhereBeneficiario = ""
  vSLocalWherePessoa = ""

  If RecordHandleOfTable("SAM_BENEFICIARIO") > 0 Then

    qBuscaContaFin.Add("SELECT HANDLE FROM SFN_CONTAFIN WHERE BENEFICIARIO = :HANDLE")
    qBuscaContaFin.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_BENEFICIARIO")
    qBuscaContaFin.Active = True
    If qBuscaContaFin.FieldByName("HANDLE").AsInteger > 0 Then
      If WebMode Then
        vSLocalWhere = " A.HANDLE "
      Else
        vSLocalWhere = " HANDLE "
      End If
      vSLocalWhere = vSLocalWhere + " IN (Select HANDLE FROM SFN_FATURA WHERE SALDO > 0 And CONTAFINANCEIRA = " + qBuscaContaFin.FieldByName("HANDLE").AsString + ")"
    End If

  ElseIf RecordHandleOfTable("SFN_PESSOA") > 0 Then

    qBuscaContaFin.Add("SELECT HANDLE FROM SFN_CONTAFIN WHERE PESSOA = :HANDLE")
    qBuscaContaFin.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SFN_PESSOA")
    qBuscaContaFin.Active = True
    If WebMode Then
      vSLocalWhere = " A.HANDLE "
    Else
      vSLocalWhere = " HANDLE "
    End If
    If qBuscaContaFin.FieldByName("HANDLE").AsInteger > 0 Then
      vSLocalWhere = vSLocalWhere + " IN (SELECT HANDLE FROM SFN_FATURA WHERE SALDO > 0 AND CONTAFINANCEIRA = " + qBuscaContaFin.FieldByName("HANDLE").AsString + ")"
    End If

  End If

  If vSLocalWhere = "" Then
    If RecordHandleOfTable("SAM_BENEFICIARIO") > 0 Then
      bsShowMessage("Beneficiário sem conta financeira", "E")
    ElseIf RecordHandleOfTable("SFN_PESSOA") > 0 Then
      bsShowMessage("Pessoa sem conta financeira", "E")
    Else
      bsShowMessage("Rotina não implementada para prestador", "E")
    End If
    CanContinue = False
  End If

  If WebMode Then
	BENEFICIARIO.WebLocalWhere = " A.HANDLE IN (SELECT BENEFICIARIO FROM SFN_CONTAFIN WHERE BENEFICIARIO IS NOT NULL)"
	PESSOA.WebLocalWhere = " A.HANDLE IN (SELECT PESSOA FROM SFN_CONTAFIN WHERE PESSOA IS NOT NULL)"
  Else
    BENEFICIARIO.LocalWhere = " HANDLE IN (SELECT BENEFICIARIO FROM SFN_CONTAFIN WHERE BENEFICIARIO IS NOT NULL)"
	PESSOA.LocalWhere = " HANDLE IN (SELECT PESSOA FROM SFN_CONTAFIN WHERE PESSOA IS NOT NULL)"
  End If

  viHContaFinOrigem = qBuscaContaFin.FieldByName("HANDLE").AsInteger
  Set qBuscaContaFin = Nothing

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim dllSamContaFinanceira As Object
  Dim qBuscaContaFinDestino As Object
  Dim vHFaturas             As String
  Dim viHContaFinDestino    As Long
  Set qBuscaContaFinDestino = NewQuery

  If (RecordHandleOfTable("SAM_BENEFICIARIO") > 0 And RecordHandleOfTable("SAM_BENEFICIARIO") = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger) Or (RecordHandleOfTable("SFN_PESSOA") > 0 And RecordHandleOfTable("SFN_PESSOA") = CurrentQuery.FieldByName("PESSOA").AsInteger) Then
      bsShowMessage("Contas de Origem e Destino não podem ser iguais !", "E")
      CanContinue = False
      Exit Sub
  End If

  qBuscaContaFinDestino.Clear
  qBuscaContaFinDestino.Add("SELECT HANDLE")
  qBuscaContaFinDestino.Add("  FROM SFN_CONTAFIN")

  If CurrentQuery.FieldByName("TABTIPORESPONSAVEL").AsInteger = 1 Then
	qBuscaContaFinDestino.Add(" WHERE BENEFICIARIO = :HANDLE ")
	qBuscaContaFinDestino.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("BENEFICIARIO").AsInteger
  ElseIf CurrentQuery.FieldByName("TABTIPORESPONSAVEL").AsInteger = 2 Then
  	qBuscaContaFinDestino.Add(" WHERE PESSOA = :HANDLE ")
	qBuscaContaFinDestino.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PESSOA").AsInteger
  End If

  qBuscaContaFinDestino.Active = True
  viHContaFinDestino = qBuscaContaFinDestino.FieldByName("HANDLE").AsInteger

  vHFaturas = formataCampoFiltroComVirgula(CurrentQuery.FieldByName("FATURAS").AsString)

  Set dllSamContaFinanceira = CreateBennerObject("SAMCONTAFINANCEIRA.TransferenciaSaldo")
  dllSamContaFinanceira.Exec(CurrentSystem, vHFaturas, viHContaFinOrigem, viHContaFinDestino)
  bsShowMessage("Transferência concluída !", "I")

  Set dllSamContaFinanceira = Nothing
  Set qBuscaContaFinDestino = Nothing
End Sub

