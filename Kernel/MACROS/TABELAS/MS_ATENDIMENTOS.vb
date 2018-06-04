'HASH: 37F89943F01C97A1C31E63A845E2F559


Public Sub FINALIZAR_OnClick()
  Dim SQL As Object
  Set SQL = NewQuery

  If Not InTransaction Then StartTransaction

  SQL.Clear
  SQL.Add("UPDATE MS_ATENDIMENTOS SET DATAFINAL = :DATAFINAL, HORAFINAL = :HORAFINAL WHERE HANDLE = :HANDLE")
  SQL.ParamByName("DATAFINAL").Value = ServerDate
  SQL.ParamByName("HORAFINAL").Value = TimeValue(Str(DatePart("h", ServerNow)) + ":" + Str(DatePart("n", ServerNow)))
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  SQL.ExecSQL

  If InTransaction Then Commit
  RefreshNodesWithTable("")

End Sub

Public Sub INICIAR_OnClick()
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Add("SELECT DATAINICIAL, HORAINICIAL FROM MS_ATENDIMENTOS WHERE EMPRESA = :EMPRESA AND FILIAL = :FILIAL AND CLINICA = :CLINICA AND DATAINICIAL IS NOT NULL AND DATAFINAL IS NULL")
  SQL.ParamByName("EMPRESA").Value = CurrentQuery.FieldByName("EMPRESA").Value
  SQL.ParamByName("FILIAL").Value = CurrentQuery.FieldByName("FILIAL").Value
  SQL.ParamByName("CLINICA").Value = CurrentQuery.FieldByName("CLINICA").Value
  SQL.Active = True
  If Not SQL.EOF Then
    MsgBox("O atendimento iniciado em " + Str(SQL.FieldByName("DATAINICIAL").AsDateTime) + " às " + Trim(Str(DatePart("h", SQL.FieldByName("HORAINICIAL").AsDateTime)) + ":" + Trim(Str(DatePart("n", SQL.FieldByName("HORAINICIAL").AsDateTime)))) + Chr(13) + "nesta mesma clínica não foi finalizado!")
    Exit Sub
  End If

  SQL.Clear
  SQL.Add("SELECT A.DATAINICIAL, A.HORAINICIAL, B.NOME PRESTNOME FROM MS_ATENDIMENTOS A, SAM_PRESTADOR B, CLI_CLINICA C WHERE A.EMPRESA = :EMPRESA AND A.FILIAL = :FILIAL AND A.CLINICA <> :CLINICA AND A.CLINICA = C.HANDLE AND C.PRESTADOR = B.HANDLE AND A.PACIENTE = :PACIENTE AND DATAINICIAL IS NOT NULL AND DATAFINAL IS NULL")
  SQL.ParamByName("EMPRESA").Value = CurrentQuery.FieldByName("EMPRESA").Value
  SQL.ParamByName("FILIAL").Value = CurrentQuery.FieldByName("FILIAL").Value
  SQL.ParamByName("CLINICA").Value = CurrentQuery.FieldByName("CLINICA").Value
  SQL.ParamByName("PACIENTE").Value = CurrentQuery.FieldByName("PACIENTE").Value
  SQL.Active = True
  If Not SQL.EOF Then
    MsgBox("Este paciente possui um atendimento iniciado em " + Str(SQL.FieldByName("DATAINICIAL").AsDateTime) + " às " + Trim(Str(DatePart("h", SQL.FieldByName("HORAINICIAL").AsDateTime)) + ":" + Trim(Str(DatePart("n", SQL.FieldByName("HORAINICIAL").AsDateTime)))) + Chr(13) + "na clínica " + SQL.FieldByName("PRESTNOME").AsString + " não finalizado.")
    Exit Sub
  End If

  If Not InTransaction Then StartTransaction

  SQL.Clear
  SQL.Add("UPDATE MS_ATENDIMENTOS SET DATAINICIAL = :DATAINICIAL, HORAINICIAL = :HORAINICIAL WHERE HANDLE = :HANDLE")
  SQL.ParamByName("DATAINICIAL").Value = ServerDate
  SQL.ParamByName("HORAINICIAL").Value = TimeValue(Str(DatePart("h", ServerNow)) + ":" + Str(DatePart("n", ServerNow)))
  SQL.ParamByName("HANDLE").Value = CurrentQuery.FieldByName("HANDLE").Value
  SQL.ExecSQL

  If InTransaction Then Commit
  RefreshNodesWithTable("")

End Sub


Public Sub TABLE_AfterScroll()
  INICIAR.Enabled = True
  FINALIZAR.Enabled = True

  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    INICIAR.Enabled = False
    FINALIZAR.Enabled = False
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("DATAINICIAL").IsNull Then
    INICIAR.Enabled = False
    Exit Sub
  End If

End Sub

