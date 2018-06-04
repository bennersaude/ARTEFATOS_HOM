'HASH: E667C116EA4EE920AE96BBE3D495E876
'#Uses "*bsShowMessage"
'CLI_INDISPONIBILIDADE

Public Sub TABLE_AfterCommitted()
  Dim BSCli001dll As Object
  Set BSCli001dll = CreateBennerObject("BSCli001.Rotinas")
  BSCli001dll.BuscaAgendaIndisponibilidade(CurrentSystem, _
                                           CurrentQuery.FieldByName("DATAHORAINICIAL").AsDateTime, _
                                           CurrentQuery.FieldByName("DATAHORAFINAL").AsDateTime, _
                                           CurrentQuery.FieldByName("RECURSO").AsInteger, _
                                           CurrentQuery.FieldByName("HANDLE").AsInteger)
  Set BSCli001dll = Nothing
End Sub

Public Sub TABLE_AfterScroll()
  'Verifica se existe alguma agena para ser remarcada
  Dim AGENDA As Object
  Set AGENDA = NewQuery
  AGENDA.Add("SELECT HANDLE FROM CLI_INDISPONIBILIDADEREMARCAR WHERE INDISPONIBILIDADE = :HANDLE")
  AGENDA.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  AGENDA.Active = True

  If Not AGENDA.EOF Then
    DATAHORAINICIAL.ReadOnly = True
    DATAHORAFINAL.ReadOnly = True
  Else
    DATAHORAINICIAL.ReadOnly = False
    DATAHORAFINAL.ReadOnly = False
  End If
  Set AGENDA = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim DATAINICIAL As Date
  Dim DATAFINAL As Date
  Dim HORAINICIAL As Date
  Dim HORAFINAL As Date
  Dim SQL As Object
  Set SQL = NewQuery

  If CurrentQuery.FieldByName("DATAHORAINICIAL").AsDateTime <ServerNow Then
    bsShowMessage("Não pode ser marcada uma indisponibilidade com início anterior ao corrente!", "E")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("DATAHORAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAHORAFINAL").AsDateTime) Then
    bsShowMessage("A data e hora inicial não pode ser superior à data e hora final!", "E")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If

  DATAINICIAL = DateValue(Str(DatePart("yyyy", CurrentQuery.FieldByName("DATAHORAINICIAL").AsDateTime)) + "-" + _
                          Str(DatePart("m", CurrentQuery.FieldByName("DATAHORAINICIAL").AsDateTime)) + "-" + _
                          Str(DatePart("d", CurrentQuery.FieldByName("DATAHORAINICIAL").AsDateTime)))
  DATAFINAL = DateValue(Str(DatePart("yyyy", CurrentQuery.FieldByName("DATAHORAFINAL").AsDateTime)) + "-" + _
                        Str(DatePart("m", CurrentQuery.FieldByName("DATAHORAFINAL").AsDateTime)) + "-" + _
                        Str(DatePart("d", CurrentQuery.FieldByName("DATAHORAFINAL").AsDateTime)))
  HORAINICIAL = CDate("1899-12-30 " + _
                Str(DatePart("h", CurrentQuery.FieldByName("DATAHORAINICIAL").AsDateTime)) + ":" + _
                Str(DatePart("n", CurrentQuery.FieldByName("DATAHORAINICIAL").AsDateTime)) + ":" + _
                Str(DatePart("n", CurrentQuery.FieldByName("DATAHORAINICIAL").AsDateTime)))
  HORAFINAL = CDate("1899-12-30 " + _
              Str(DatePart("h", CurrentQuery.FieldByName("DATAHORAFINAL").AsDateTime)) + ":" + _
              Str(DatePart("n", CurrentQuery.FieldByName("DATAHORAFINAL").AsDateTime)) + ":" + _
              Str(DatePart("n", CurrentQuery.FieldByName("DATAHORAFINAL").AsDateTime)))
  SQL.Active = False
  SQL.Clear
  If FormatDateTime2("DD/MM/YYYY", CurrentQuery.FieldByName("DATAHORAINICIAL").AsDateTime) = FormatDateTime2("DD/MM/YYYY", CurrentQuery.FieldByName("DATAHORAFINAL").AsDateTime)Then
    SQL.Add("SELECT HANDLE")
    SQL.Add("  FROM CLI_AGENDA")
    SQL.Add(" WHERE DATAMARCADA = :DATAINICIAL")
    SQL.Add("   AND HORAMARCADA >= :HORAINICIAL")
    SQL.Add("   AND HORAMARCADA <= :HORAFINAL")
    SQL.Add("   AND HORADESMARCACAO IS NULL")
    SQL.Add("   AND RECURSO = :RECURSO")
  Else
    SQL.Add("SELECT HANDLE   ")
    SQL.Add("  FROM CLI_AGENDA")
    SQL.Add(" WHERE (   ")
    SQL.Add("        (DATAMARCADA = :DATAINICIAL AND HORAMARCADA >= :HORAINICIAL)   ")
    SQL.Add("    OR  (DATAMARCADA = :DATAFINAL   AND HORAMARCADA <= :HORAFINAL)   ")
    SQL.Add("    OR  (DATAMARCADA > :DATAINICIAL AND DATAMARCADA <  :DATAFINAL)   ")
    SQL.Add("       )   ")
    SQL.Add("   AND DATADESMARCACAO IS NULL    ")
    SQL.Add("   AND RECURSO = :RECURSO   ")
    SQL.ParamByName("DATAFINAL").AsDateTime = DATAFINAL
  End If
  SQL.ParamByName("RECURSO").AsInteger = CurrentQuery.FieldByName("RECURSO").AsInteger
  SQL.ParamByName("DATAINICIAL").AsDateTime = DATAINICIAL
  SQL.ParamByName("HORAINICIAL").AsDateTime = HORAINICIAL
  SQL.ParamByName("HORAFINAL").AsDateTime = HORAFINAL
  SQL.Active = True
  If Not SQL.EOF Then
    If bsShowMessage("Já existem consultas marcadas para esse período! Continuar?", "Q") = vbNo Then
      CanContinue = False
      Set SQL = Nothing
      Exit Sub
    End If
  End If

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT 1    																	")
  SQL.Add("FROM CLI_INDISPONIBILIDADE I    												")
  SQL.Add("WHERE I.RECURSO = :RECURSO AND   											")
  SQL.Add("    (   																		")
  SQL.Add("    ((DATAHORAINICIAL >= :DATAINICIAL) AND (DATAHORAFINAL <= :DATAINICIAL)   ")
  SQL.Add("      OR    																	")
  SQL.Add("      (DATAHORAINICIAL >= :DATAFINAL) AND (DATAHORAFINAL <= :DATAFINAL)) 	")
  SQL.Add("OR    																		")
  SQL.Add("      ((DATAHORAINICIAL >= :DATAINICIAL) AND (DATAHORAINICIAL <= :DATAFINAL) ")
  SQL.Add("      OR    																	")
  SQL.Add("      (DATAHORAFINAL >= :DATAINICIAL) AND (DATAHORAFINAL <= :DATAFINAL))    	")
  SQL.Add("      )    																	")

  If CurrentQuery.FieldByName("HANDLE").AsInteger > 0 Then
    SQL.Add("AND I.HANDLE <> " + CurrentQuery.FieldByName("HANDLE").AsString )
  End If

  SQL.ParamByName("RECURSO").AsInteger = CurrentQuery.FieldByName("RECURSO").AsInteger
  SQL.ParamByName("DATAINICIAL").AsDateTime = CurrentQuery.FieldByName("DATAHORAINICIAL").AsDateTime
  SQL.ParamByName("DATAFINAL").AsDateTime = CurrentQuery.FieldByName("DATAHORAFINAL").AsDateTime
  SQL.Active = True

  If Not SQL.EOF Then
    MsgBox("Já existe indisponibilidade cadastrada que coincide com este período!")
    CanContinue = False
    Set SQL = Nothing
    Exit Sub
  End If

  Set SQL = Nothing
End Sub

