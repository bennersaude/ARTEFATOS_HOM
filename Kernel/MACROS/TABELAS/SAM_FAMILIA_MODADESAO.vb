'HASH: D1DF4120E84CC9514769820409CCFC19
'Macro: SAM_FAMILIA_MODADESAO
'#Uses "*bsShowMessage"

Dim vDATAULTIMOREAJUSTE As Date

Public Sub DATAADESAO_OnExit()
  If CurrentQuery.State = 3 Or (CurrentQuery.State = 2 And CurrentQuery.FieldByName("DATAADESAO").AsDateTime <> vDATAULTIMOREAJUSTE) Then
    If Not CurrentQuery.FieldByName("DATAADESAO").IsNull Then
      CurrentQuery.FieldByName("DATAULTIMOREAJUSTE").Value = CurrentQuery.FieldByName("DATAADESAO").AsDateTime
    End If
  End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
  vDATAULTIMOREAJUSTE = CurrentQuery.FieldByName("DATAADESAO").Value
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)

  Dim SQL
  Set SQL = NewQuery
  SQL.Add("SELECT DATACANCELAMENTO FROM SAM_FAMILIA_MOD WHERE HANDLE = :HFAMILIAMOD")
  SQL.ParamByName("HFAMILIAMOD").Value = RecordHandleOfTable("SAM_FAMILIA_MOD")
  SQL.Active = True
  If Not SQL.FieldByName("DATACANCELAMENTO").IsNull Then
    bsShowMessage("Módulo cancelado não permite manutenções", "I")
    CurrentQuery.Cancel
    RefreshNodesWithTable("SAM_FAMILIA_MODADESAO")
  End If

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qFechamento As Object
  Set qFechamento = NewQuery

  qFechamento.Add("SELECT DATAFECHAMENTO FROM SAM_PARAMETROSBENEFICIARIO")
  qFechamento.Active = True

  If CurrentQuery.State = 3 Then

    If CurrentQuery.FieldByName("DATAADESAO").AsDateTime < qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime Then
      CanContinue = False
      bsShowMessage("Não é possível cadastrar data de adesão inferior a data de fechamento - Parâmetros Gerais", "E")
    End If
    If Not CurrentQuery.FieldByName("DATAULTIMOREAJUSTE").IsNull Then
      If CurrentQuery.FieldByName("DATAULTIMOREAJUSTE").AsDateTime < qFechamento.FieldByName("DATAFECHAMENTO").AsDateTime Then
        CanContinue = False
        bsShowMessage("Não é possível cadastrar data do último reajuste inferior a data de fechamento - Parâmetros Gerais", "E")
      End If
    End If
  End If

  Set qFechamento = Nothing

End Sub

