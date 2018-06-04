'HASH: E2484F668A19D167099980DB9DB5E5E1
'#Uses "*bsShowMessage"
Public Sub PLANOLIMITACAO_OnPopup(ShowPopup As Boolean)
  Dim Procura As Object
  Dim handlexx As Long

  ShowPopup = False
  Set Procura = CreateBennerObject("Procura.Procurar")
  handlexx = Procura.Exec(CurrentSystem, "SAM_PLANO_LIMITACAO|SAM_LIMITACAO[SAM_PLANO_LIMITACAO.LIMITACAO = SAM_LIMITACAO.HANDLE]", "DESCRICAO", 1, "Descrição", "PLANO = " + Str(RecordHandleOfTable("SAM_PLANO")), "Procura por Limitação", True, "")
  If handlexx <>0 Then
    CurrentQuery.Edit
    CurrentQuery.FieldByName("PLANOLIMITACAO").Value = handlexx

    'SMS 61198 - Matheus - Início
    Dim sql As Object
    Set sql = NewQuery

    sql.Active = False
    sql.Clear
    sql.Add("SELECT SL.PERIODICIDADE          ")
    sql.Add("  FROM SAM_LIMITACAO SL,         ")
    sql.Add("       SAM_PLANO_LIMITACAO SPL   ")
    sql.Add(" WHERE SPL.HANDLE = :HANDLE      ")
    sql.Add("   AND SL.HANDLE = SPL.LIMITACAO ")
    sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PLANOLIMITACAO").AsInteger
    sql.Active = True

    If sql.FieldByName("PERIODICIDADE").AsInteger = 2 Then
      CurrentQuery.FieldByName("PERIODO").AsInteger = 1
      PERIODO.Visible = False
    Else
      CurrentQuery.FieldByName("PERIODO").Clear
      PERIODO.Visible = True
    End If

    Set sql = Nothing
    'SMS 61198 - Matheus - Fim
  End If
  Set Procura = Nothing

End Sub

Public Sub TABLE_AfterScroll()
 'SMS 61198 - Matheus - Início
  Dim sql As Object
  Set sql = NewQuery

  sql.Active = False
  sql.Clear
  sql.Add("SELECT SL.PERIODICIDADE          ")
  sql.Add("  FROM SAM_LIMITACAO SL,         ")
  sql.Add("       SAM_PLANO_LIMITACAO SPL   ")
  sql.Add(" WHERE SPL.HANDLE = :HANDLE      ")
  sql.Add("   AND SL.HANDLE = SPL.LIMITACAO ")
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PLANOLIMITACAO").AsInteger
  sql.Active = True

  If sql.FieldByName("PERIODICIDADE").AsInteger = 2 Then
    PERIODO.Visible = False
  Else
    PERIODO.Visible = True
  End If

  Set sql = Nothing
  'SMS 61198 - Matheus - Fim
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'SMS 61198 - Matheus - Início
  Dim sql As Object
  Set sql = NewQuery

  sql.Active = False
  sql.Clear
  sql.Add("SELECT SL.PERIODICIDADE          ")
  sql.Add("  FROM SAM_LIMITACAO SL,         ")
  sql.Add("       SAM_PLANO_LIMITACAO SPL   ")
  sql.Add(" WHERE SPL.HANDLE = :HANDLE      ")
  sql.Add("   AND SL.HANDLE = SPL.LIMITACAO ")
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("PLANOLIMITACAO").AsInteger
  sql.Active = True

  If sql.FieldByName("PERIODICIDADE").AsInteger = 2 Then  CurrentQuery.FieldByName("PERIODO").AsInteger = 1

  Set sql = Nothing
  'SMS 61198 - Matheus - Fim
End Sub
