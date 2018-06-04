'HASH: 4EB548E62E5BDDE11AFA09B38BBC369A
'Macro: SAM_MIGRACAO_PROCESSOBENEF
'#Uses "*bsShowMessage"


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim vFrase As String
  Dim SQL As Object

  Set SQL = NewQuery

  vFrase = "SELECT RESPONSAVEL FROM SAM_MIGRACAO_PROCESSO WHERE HANDLE = :vHandle"
  SQL.Clear
  SQL.Add(vFrase)
  SQL.ParamByName("vHandle").Value = CurrentQuery.FieldByName("MIGRACAOPROCESSO").AsInteger
  SQL.Active = True

  If CurrentUser = SQL.FieldByName("RESPONSAVEL").AsInteger Then

    SQL.Active = False

    vFrase = "DELETE FROM SAM_MIGRACAO_PROCESSOBENEF_CAR WHERE SAM_MIGRACAO_PROCESSOBENEF_CAR.MIGRACAOPROCESSOBENEF = :vHandle"
    SQL.Clear
    SQL.Add(vFrase)
    SQL.ParamByName("vHandle").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL

    vFrase = "DELETE FROM SAM_MIGRACAO_PROCESSOBENEF_LIM  WHERE SAM_MIGRACAO_PROCESSOBENEF_LIM.MIGRACAOPROCESSOBENEF = :vHandle"
    SQL.Clear
    SQL.Add(vFrase)
    SQL.ParamByName("vHandle").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL

    vFrase = "DELETE FROM SAM_MIGRACAO_PROCESSOBENEF_MOD WHERE SAM_MIGRACAO_PROCESSOBENEF_MOD.MIGRACAOPROCESSOBENEF = :vHandle"
    SQL.Clear
    SQL.Add(vFrase)
    SQL.ParamByName("vHandle").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ExecSQL


  Else

    CanContinue = False
    bsShowMessage("Operação cancelada. Usuário não é o Responsável", "E")

  End If

  Set SQL = Nothing

End Sub


