'HASH: 6858E72788F50EC0B5EA62296DB5D75F
'CLI_ESCALA SHIBA -03/12/2001
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
  DIASEMANASTR.Visible = False
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Ok As Boolean
  Dim SQL As Object

  Set SQL = NewQuery
  SQL.Clear
  SQL.Add("SELECT CLINICA FROM CLI_RECURSO WHERE HANDLE = :RECURSO")
  SQL.ParamByName("RECURSO").AsInteger = CurrentQuery.FieldByName("RECURSO").AsInteger

  Select Case CurrentQuery.FieldByName("DIASEMANA").AsInteger
	Case 1
		CurrentQuery.FieldByName("DIASEMANASTR").AsString = "DOMINGO"
	Case 2
		CurrentQuery.FieldByName("DIASEMANASTR").AsString = "SEGUNDA"
	Case 3
		CurrentQuery.FieldByName("DIASEMANASTR").AsString = "TERÇA"
	Case 4
		CurrentQuery.FieldByName("DIASEMANASTR").AsString = "QUARTA"
	Case 5
		CurrentQuery.FieldByName("DIASEMANASTR").AsString = "QUINTA"
	Case 6
		CurrentQuery.FieldByName("DIASEMANASTR").AsString = "SEXTA"
	Case 7
		CurrentQuery.FieldByName("DIASEMANASTR").AsString = "SÁBADO"
	Case 8
		CurrentQuery.FieldByName("DIASEMANASTR").AsString = "SEGUNDA À SEXTA"
	Case 9
		CurrentQuery.FieldByName("DIASEMANASTR").AsString = "SEGUNDA À SABADO"
	Case 10
		CurrentQuery.FieldByName("DIASEMANASTR").AsString = "SEGUNDA À DOMINGO"
  End Select

  SQL.Active = True
  Dim ClinicaDLL As Object
  Set ClinicaDLL = CreateBennerObject("CliClinica.Agenda")
  ClinicaDLL.EscalaSobreposta(CurrentSystem, CurrentQuery.FieldByName("DATAINICIAL").AsDateTime, _
                              CurrentQuery.FieldByName("DATAFINAL").AsDateTime, _
                              CurrentQuery.FieldByName("HORAINICIALINTERVALO").AsDateTime, _
                              CurrentQuery.FieldByName("HORAFINALINTERVALO").AsDateTime, _
                              CurrentQuery.FieldByName("HORAINICIAL").AsDateTime, _
                              CurrentQuery.FieldByName("HORAFINAL").AsDateTime, _
                              CurrentQuery.FieldByName("RECURSO").AsInteger, _
                              CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger, _
                              CurrentQuery.FieldByName("HANDLE").AsInteger, _
                              CurrentQuery.FieldByName("DIASEMANA").AsInteger, _
                              SQL.FieldByName("CLINICA").AsInteger, _
                              Ok)
  If Not Ok Then
    CanContinue = False
    Exit Sub
  End If
  Set ClinicaDLL = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim qVerifica As Object
  Dim vDiaSemana As String

  Set qVerifica = NewQuery

  qVerifica.Add("SELECT COUNT(*) TOTAL FROM CLI_AGENDA")
  qVerifica.Add(" WHERE RECURSO = :RECURSO")
  qVerifica.Add("   AND ESPECIALIDADE = :ESPECIALIDADE")
  qVerifica.Add("   AND DATAMARCADA >= :DATAINICIAL")
  qVerifica.Add("   AND HORAMARCADA BETWEEN :HORAINICIAL AND :HORAFINAL")
  qVerifica.Add("   AND MOTIVODESMARCACAO IS NULL")

  If CurrentQuery.FieldByName("DIASEMANA").AsInteger <> 10 Then
    If CurrentQuery.FieldByName("DIASEMANA").AsInteger = 8 Then
      vDiaSemana = "(2,3,4,5,6)"
    ElseIf CurrentQuery.FieldByName("DIASEMANA").AsInteger = 9 Then
      vDiaSemana = "(2,3,4,5,6,7)"
    Else
      vDiaSemana = "(" + CurrentQuery.FieldByName("DIASEMANA").AsString + ")"
    End If
    If InStr(SQLServer, "MSSQL") > 0 Then
      qVerifica.Add("   AND DATEPART(DW,DATAMARCADA) IN " + vDiaSemana)
    ElseIf InStr(SQLServer, "DB2") > 0 Then
      qVerifica.Add("   AND DAYOFWEEK(DATAMARCADA) IN " + vDiaSemana)
    ElseIf InStr(SQLServer, "ORA") > 0 Then
      qVerifica.Add("   AND TO_CHAR(DATAMARCADA, 'd') IN " + vDiaSemana)
    ElseIf InStr(SQLServer, "CACHE")>0 Then
      qVerifica.Add("   AND { fn DAYOFWEEK(DATAMARCADA) } IN " + vDiaSemana)
    End If
  End If

  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    qVerifica.Add("   AND DATAMARCADA <= :DATAFINAL")
    qVerifica.ParamByName("DATAFINAL").Value = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
  End If

  qVerifica.ParamByName("RECURSO").Value = CurrentQuery.FieldByName("RECURSO").AsInteger
  qVerifica.ParamByName("ESPECIALIDADE").Value = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
  qVerifica.ParamByName("DATAINICIAL").Value = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
  qVerifica.ParamByName("HORAINICIAL").Value = CurrentQuery.FieldByName("HORAINICIAL").AsDateTime
  qVerifica.ParamByName("HORAFINAL").Value = CurrentQuery.FieldByName("HORAFINAL").AsDateTime
  qVerifica.Active = True

  If qVerifica.FieldByName("TOTAL").AsInteger > 0 Then
    bsShowMessage("Existem consultas agendadas!", "E")
    CanContinue = False
  End If

  Set qVerifica = Nothing
End Sub

