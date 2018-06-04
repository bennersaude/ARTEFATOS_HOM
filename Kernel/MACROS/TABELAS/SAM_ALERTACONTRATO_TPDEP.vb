'HASH: 538F49F592EB9163E8C11A7995BEC076
'Macro: SAM_ALERTACONTRATO_TPDEP
'#Uses "*bsShowMessage"

Public Sub CONTRATOTIPODEPENDENTE_OnPopup(ShowPopup As Boolean)
  Dim ProcuraDLL As Variant
  Dim vColunas As String
  Dim vCampos As String
  Dim vCriterio As String
  Dim vHandle As Long
  Dim vUsuario As String
  Dim qContrato As BPesquisa

  Set qContrato = NewQuery
  qContrato.Clear
  qContrato.Active = False
  qContrato.Add("SELECT AC.CONTRATO              ")
  qContrato.Add("  FROM SAM_ALERTACONTRATO AC    ")
  qContrato.Add(" WHERE HANDLE = :HANDLECONTRATO ")
  qContrato.ParamByName("HANDLECONTRATO").AsInteger = CInt(RecordHandleOfTable("SAM_ALERTACONTRATO"))
  qContrato.Active = True

  vUsuario = Str(CurrentUser)
  Set ProcuraDLL = CreateBennerObject("PROCURA.PROCURAR")

  vColunas = "SAM_TIPODEPENDENTE.DESCRICAO"
  vCriterio = "CONTRATO = " + qContrato.FieldByName("CONTRATO").AsString
  vCampos = "Descrição"
  vHandle = ProcuraDLL.Exec(CurrentSystem, "SAM_CONTRATO_TPDEP|SAM_TIPODEPENDENTE[SAM_CONTRATO_TPDEP.TIPODEPENDENTE=SAM_TIPODEPENDENTE.HANDLE]", vColunas, 2, vCampos, vCriterio, "Tipo de dependente", True, "")
  ShowPopup = False

  If vHandle <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("CONTRATOTIPODEPENDENTE").Value = vHandle
  End If

  ShowPopup = False
  Set ProcuraDLL = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  	If WebMode Then
      Dim SQL As Object

      Set SQL = NewQuery
      SQL.Add("SELECT CONTRATO")
      SQL.Add("FROM SAM_ALERTACONTRATO")
      SQL.Add("WHERE HANDLE = :HALERTACONTRATO")
      SQL.ParamByName("HALERTACONTRATO").AsInteger = CurrentQuery.FieldByName("ALERTACONTRATO").AsInteger
      SQL.Active = True

   	  CONTRATOTIPODEPENDENTE.WebLocalWhere = "A.CONTRATO = " + SQL.FieldByName("CONTRATO").AsString

   	  Set SQL = Nothing
	End If

  Dim Q As Object

  Set Q = NewQuery
  Q.Add("SELECT * FROM SAM_ALERTACONTRATO WHERE HANDLE = :CONTRATOALERTA")
  Q.ParamByName("CONTRATOALERTA").Value = CurrentQuery.FieldByName("ALERTACONTRATO").AsInteger
  Q.Active = True

  CurrentQuery.RequestLive = Q.FieldByName("DATAFINAL").IsNull
  CONTRATOTIPODEPENDENTE.ReadOnly = Not Q.FieldByName("DATAFINAL").IsNull
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Q As Object

  Set Q = NewQuery
  Q.Add("SELECT * FROM SAM_ALERTACONTRATO WHERE HANDLE = :HANDLE")
  Q.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("ALERTACONTRATO").AsInteger
  Q.Active = True

  If Not Q.FieldByName("DATAFINAL").IsNull Then
    bsShowMessage("Cadastro não permitido, pois a vigência do alerta está fechada", "E")
    CanContinue = False
    Exit Sub
  End If
End Sub
