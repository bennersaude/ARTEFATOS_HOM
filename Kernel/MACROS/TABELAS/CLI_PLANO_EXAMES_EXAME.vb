'HASH: CC7BBDD7AE53C04CA8573690AFDD632A
 '#Uses "*bsShowMessage"
 Option Explicit

Dim viEstadoQuery As Long

Public Sub TABLE_AfterDelete()
  ExcluirDoResultado
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  ExcluirDoResultado
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim sql As Object

  viEstadoQuery = CurrentQuery.State

  Set sql = NewQuery

  sql.Clear
  sql.Add("SELECT DESCRICAO")
  sql.Add("  FROM SAM_TGE  ")
  sql.Add(" WHERE HANDLE = :HANDLE")
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  sql.Active =  True

  CurrentQuery.FieldByName("DESCRICAO").AsString = sql.FieldByName("DESCRICAO").AsString

  Set sql = Nothing

End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
  If CommandID = "MONITORAMENTO" Then
    If IncluirNoMonitoramento Then
      bsShowMessage("Exame incluído no monitoramento!","I")
    Else
      bsShowMessage("Não foi possível incluir no monitoramento, o procedimento não consta cadastro geral de monitoramento!","I")
    End If
  End If
End Sub
Function IncluirNoMonitoramento As Boolean
  Dim sqlMonitoramento As Object
  Dim sqlPaciente As Object
  Dim sqlInsert As Object

  Set sqlPaciente = NewQuery

  sqlPaciente.Clear
  sqlPaciente.Add("SELECT S.PACIENTE                                   ")
  sqlPaciente.Add("  FROM CLI_SUBJETIVO S                              ")
  sqlPaciente.Add("  JOIN CLI_PLANO_EXAMES E ON E.SUBJETIVO = S.HANDLE ")
  sqlPaciente.Add(" WHERE E.HANDLE = :HANDLESOLICIT                    ")
  sqlPaciente.ParamByName("HANDLESOLICIT").AsInteger = CurrentQuery.FieldByName("EXAMESOLICIT").AsInteger
  sqlPaciente.Active = True

  Set sqlMonitoramento = NewQuery

  sqlMonitoramento.Clear
  sqlMonitoramento.Add("SELECT HANDLE")
  sqlMonitoramento.Add("  FROM CLI_MONITORAMENTO ")
  sqlMonitoramento.Add(" WHERE PROCEDIMENTO = :EVENTO ")
  sqlMonitoramento.Add("   AND MONITORAMENTOGERAL = 'N' ")
  sqlMonitoramento.ParamByName("EVENTO").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
  sqlMonitoramento.Active = True

  If Not sqlMonitoramento.EOF Then

    If Not ConstaNoMonitoramentoEspecifico(sqlPaciente.FieldByName("PACIENTE").AsInteger, sqlMonitoramento.FieldByName("HANDLE").AsInteger) Then

      Set sqlInsert = NewQuery

      sqlInsert.Clear
      sqlInsert.Add("INSERT INTO CLI_PACIENTE_MONITORAMENTO           ")
      sqlInsert.Add("            (HANDLE, PACIENTE, MONITORAMENTO)    ")
      sqlInsert.Add("     VALUES (:HANDLE, :PACIENTE, :MONITORAMENTO) ")
      sqlInsert.ParamByName("HANDLE").AsInteger = NewHandle("CLI_PACIENTE_MONITORAMENTO")
      sqlInsert.ParamByName("PACIENTE").AsInteger = sqlPaciente.FieldByName("PACIENTE").AsInteger
      sqlInsert.ParamByName("MONITORAMENTO").AsInteger = sqlMonitoramento.FieldByName("HANDLE").AsInteger
      sqlInsert.ExecSQL

      Set sqlInsert = Nothing

	  IncluirNoMonitoramento = True

    End If

  Else
    IncluirNoMonitoramento = False
  End If

  Set sqlPaciente = Nothing
  Set sqlMonitoramento = Nothing

End Function

Function ExcluirDoResultado As Boolean

  Dim sql As Object

  Set sql = NewQuery

  sql.Clear
  sql.Add("DELETE FROM CLI_EXAME_RESULTADO ")
  sql.Add(" WHERE EXAME = :EXAME           ")
  sql.ParamByName("EXAME").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.ExecSQL

  Set sql = Nothing

End Function
Function ConstaNoMonitoramentoEspecifico(paciente As Long, monitoramento As Long) As Boolean

  Dim sql As Object

  Set sql = NewQuery

  sql.Clear
  sql.Add("SELECT HANDLE                         ")
  sql.Add("  FROM CLI_PACIENTE_MONITORAMENTO     ")
  sql.Add(" WHERE PACIENTE = :PACIENTE           ")
  sql.Add("   AND MONITORAMENTO = :MONITORAMENTO ")
  sql.ParamByName("PACIENTE").AsInteger = paciente
  sql.ParamByName("MONITORAMENTO").AsInteger = monitoramento
  sql.Active = True

  If sql.EOF Then
    ConstaNoMonitoramentoEspecifico = False
  Else
    ConstaNoMonitoramentoEspecifico = True
  End If

  Set sql = Nothing

End Function
