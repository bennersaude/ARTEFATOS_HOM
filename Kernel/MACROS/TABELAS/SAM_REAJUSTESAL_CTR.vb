'HASH: 5A2BD057662270E4975B53E73F77AFF2
'#Uses "*bsShowMessage"

Dim viLotacoes As Long
Dim vGerado As Boolean

Public Sub CONTRATO_OnPopup(ShowPopup As Boolean)
  Dim interface As Object
  Dim vHandle As Long
  Dim vCampos As String
  Dim vColunas As String
  Dim vCriterio As String

  ShowPopup = False
  Set interface = CreateBennerObject("Procura.Procurar")

  vColunas = "CONTRATO|CONTRATANTE|DATAADESAO"
  vCriterio = "DATACANCELAMENTO IS NULL "
  vCampos = "Nº do Contrato|Contratante|Data Adesão"

  vHandle = interface.Exec(CurrentSystem, "SAM_CONTRATO", vColunas, 1, vCampos, vCriterio, "Contratos", False, CONTRATO.LocateText)

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATO").Value = vHandle
  End If

  Set interface = Nothing
End Sub

Public Sub TABLE_AfterScroll()
	Dim Query As Object
	Set Query = NewQuery

	Query.Clear
	Query.Add("SELECT SITUACAOGERACAO")
	Query.Add("  FROM SAM_REAJUSTESAL_PARAM")
	Query.Add(" WHERE HANDLE = :HANDLE")
	Query.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_REAJUSTESAL_PARAM")
	Query.Active = True

	vGerado = Query.FieldByName("SITUACAOGERACAO").AsInteger <> 1

	Query.Clear
	Query.Add("SELECT COUNT(HANDLE) QTD")
	Query.Add("  FROM SAM_REAJUSTESAL_CTR_LOTACAO")
	Query.Add(" WHERE REAJUSTESALCTR = :HANDLE")
	Query.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	Query.Active = True

	viLotacoes = Query.FieldByName("QTD").AsInteger

	CONTRATO.ReadOnly = viLotacoes > 0
	Set Query = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)

	If vGerado Then
		bsShowMessage("O contrato não pode ser excluído pois a rotina já está gerada!","E")
		CanContinue = False
	End If

	If viLotacoes > 0 Then
		bsShowMessage("Este contrato não pode ser excluído pois existem lotações cadastradas!","E")
		CanContinue = False
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	If vGerado Then
		bsShowMessage("Não é permitido incluir contratos pois a rotina já está gerada!","E")
		CanContinue = False
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Consulta As Object
  Set Consulta = NewQuery

  Consulta.Active = False
  Consulta.Clear
  Consulta.Add("SELECT S.CONTRATO                                       ")
  Consulta.Add("  FROM SAM_REAJUSTESAL_CTR S                            ")
  Consulta.Add("  JOIN SAM_REAJUSTESAL_PARAM R ON R.HANDLE = S.REAJUSTESAL")
  Consulta.Add(" WHERE S.CONTRATO = :CONTRATO                           ")
  Consulta.Add("   AND R.HANDLE = :ROTINA                               ")
  If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
    Consulta.Add(" AND S.HANDLE <> :HANDLE   ")
    Consulta.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  End If

  Consulta.ParamByName("CONTRATO").AsInteger = CurrentQuery.FieldByName("CONTRATO").AsInteger
  Consulta.ParamByName("ROTINA").AsInteger = CurrentQuery.FieldByName("REAJUSTESAL").AsInteger
  Consulta.Active = True

  If Not Consulta.FieldByName("CONTRATO").IsNull Then
    bsShowMessage("Contrato já cadastrado na rotina!","E")
    CanContinue = False
    Exit Sub
  End If
End Sub

Public Sub TABLE_OnInsertBtnClick(CanContinue As Boolean)

	Dim Query As Object
	Set Query = NewQuery

	Query.Clear
	Query.Add("SELECT SITUACAOGERACAO")
	Query.Add("  FROM SAM_REAJUSTESAL_PARAM")
	Query.Add(" WHERE HANDLE = :HANDLE")
	Query.ParamByName("HANDLE").AsInteger = RecordHandleOfTable("SAM_REAJUSTESAL_PARAM")
	Query.Active = True

	vGerado = Query.FieldByName("SITUACAOGERACAO").AsInteger <> 1

	If vGerado Then
		bsShowMessage("Não é permitido incluir contratos pois a rotina já está gerada!","E")
		CanContinue = False
	End If
End Sub
