'HASH: 06C2BFBC23174510180AB6494203503E
'SFN_IRRF_AGENDA
'#Uses "*bsShowMessage"


Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Clear
  SQL.Active = False
  If CurrentQuery.State = 2 Then ' Alteração
    SQL.Add("SELECT HANDLE FROM SFN_IRRF_AGENDA WHERE HANDLE<>:HAGENDA AND DATAAGENDA=:DATAAGENDA AND TIPO=:HTIPO")
    SQL.ParamByName("HAGENDA").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ParamByName("HTIPO").AsString = CurrentQuery.FieldByName("TIPO").AsString
  Else
    SQL.Add("SELECT HANDLE FROM SFN_IRRF_AGENDA WHERE DATAAGENDA=:DATAAGENDA AND TIPO=:HTIPO")
    SQL.ParamByName("HTIPO").AsString = CurrentQuery.FieldByName("TIPO").AsString
  End If
  SQL.ParamByName("DATAAGENDA").AsDateTime = CurrentQuery.FieldByName("DATAAGENDA").AsDateTime
  SQL.Active = True

  If Not SQL.EOF Then
    bsShowMessage("Já existe um registro com esta mesma Data de Agendamento", "E")
    CanContinue = False
    Exit Sub
  End If

  If (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime > CurrentQuery.FieldByName("DATAFINAL").AsDateTime) Or _
       (CurrentQuery.FieldByName("DATAAGENDA").AsDateTime > CurrentQuery.FieldByName("DATARECOLHIMENTO").AsDateTime) Or _
       (CurrentQuery.FieldByName("DATAFINAL").AsDateTime > CurrentQuery.FieldByName("DATAAGENDA").AsDateTime) Then
    bsShowMessage("Datas de Agendamento incorreta ! Verifique estes motivos:" + Chr(13) + Chr(13) + "Data de Agendamento maior que a Data de Recolhimento ou" + Chr(13) + _
           "Data Inicial maior que a Data Final ou" + Chr(13) + "Data Final maior que a Data Agenda", "E")
    CanContinue = False
    Exit Sub
  End If

  SQL.Clear
  SQL.Active = False
  If CurrentQuery.State = 2 Then ' Alteração
    SQL.Add("SELECT DATAINICIAL, DATAFINAL FROM SFN_IRRF_AGENDA WHERE HANDLE<>:HANDLE AND TIPO=:HTIPO")
    SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
    SQL.ParamByName("HTIPO").AsString = CurrentQuery.FieldByName("TIPO").AsString
  Else
    SQL.Add("SELECT DATAINICIAL, DATAFINAL FROM SFN_IRRF_AGENDA WHERE TIPO=:HTIPO")
    SQL.ParamByName("HTIPO").AsString = CurrentQuery.FieldByName("TIPO").AsString
  End If
  SQL.Active = True
  SQL.First
  While Not SQL.EOF
    If (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime >= SQL.FieldByName("DATAINICIAL").AsDateTime) And _
        (CurrentQuery.FieldByName("DATAINICIAL").AsDateTime <= SQL.FieldByName("DATAFINAL").AsDateTime) Then
      bsShowMessage("     Data Inicial do Agendamento inválida!" + Chr(13) + _
             "Verifique se ela está dentro de alguma vigência", "E")
      CanContinue = False
      Exit Sub
    Else
      If (CurrentQuery.FieldByName("DATAFINAL").AsDateTime >= SQL.FieldByName("DATAINICIAL").AsDateTime) And _
          (CurrentQuery.FieldByName("DATAFINAL").AsDateTime <= SQL.FieldByName("DATAFINAL").AsDateTime) Then
        bsShowMessage("     Data Final do Agendamento inválida!" + Chr(13) + _
               "Verifique se ela está dentro de alguma vigência", "E")
        CanContinue = False
        Exit Sub
      End If
    End If
    SQL.Next
  Wend
  Set SQL = Nothing
End Sub

