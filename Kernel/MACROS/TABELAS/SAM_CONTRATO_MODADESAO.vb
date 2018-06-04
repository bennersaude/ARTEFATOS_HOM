'HASH: A83B63CEF029AAB175EFD58245F48464
'Macro: SAM_CONTRATO_MODADESAO
'#Uses "*bsShowMessage"

Dim vDATAULTIMOREAJUSTE As Date

Public Sub DATAADESAO_OnExit()
  If CurrentQuery.State = 3 Or (CurrentQuery.State = 2 And CurrentQuery.FieldByName("DATAADESAO").AsDateTime <> vDATAULTIMOREAJUSTE) Then
    CurrentQuery.FieldByName("DATAULTIMOREAJUSTE").Value = CurrentQuery.FieldByName("DATAADESAO").Value
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  vDATAULTIMOREAJUSTE = CurrentQuery.FieldByName("DATAADESAO").Value
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  Dim SQL
  Set SQL = NewQuery
  SQL.Add("SELECT DATACANCELAMENTO FROM SAM_CONTRATO_MOD WHERE HANDLE = :HCONTRATOMOD")
  SQL.ParamByName("HCONTRATOMOD").Value = RecordHandleOfTable("SAM_CONTRATO_MOD")
  SQL.Active = True
  If Not SQL.FieldByName("DATACANCELAMENTO").IsNull Then
    bsShowMessage("Módulo cancelado não permite manutenções", "I")
    CurrentQuery.Cancel
    RefreshNodesWithTable("SAM_CONTRATO_MODADESAO")
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  If CurrentQuery.State = 3 Or (CurrentQuery.State = 2 And CurrentQuery.FieldByName("DATAADESAO").AsDateTime <> vDATAULTIMOREAJUSTE) Then
    CurrentQuery.FieldByName("DATAULTIMOREAJUSTE").Value = CurrentQuery.FieldByName("DATAADESAO").Value
  End If


  Dim SQLFechamento
  Set SQLFechamento = NewQuery
  SQLFechamento.Add("SELECT DATAFECHAMENTO FROM SAM_PARAMETROSBENEFICIARIO")
  SQLFechamento.Active = True

  If CurrentQuery.State = 3 Then
    If SQLFechamento.FieldByName("DATAFECHAMENTO").AsDateTime > CurrentQuery.FieldByName("DATAADESAO").AsDateTime Then
      bsShowMessage("Não é possível cadastrar data inicial inferior a data de fechamento - Parâmetros Gerais", "E")
      CanContinue = False
    End If

    If SQLFechamento.FieldByName("DATAFECHAMENTO").AsDateTime > CurrentQuery.FieldByName("DATAULTIMOREAJUSTE").AsDateTime Then
      bsShowMessage("Não é possível cadastrar data final inferior a data de fechamento - Parâmetros Gerais", "E")
      CanContinue = False
    End If
  End If

  Set SQLFechamento = Nothing

End Sub

