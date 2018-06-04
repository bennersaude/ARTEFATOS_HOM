'HASH: AB90169C2ABCF6F7EF202D6BEAE72184
'Macro: SAM_PROJECAOVENCIMENTO
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Add("SELECT HANDLE")
  SQL.Add("FROM SAM_PROJECAOVENCIMENTO")
  SQL.Add("WHERE HANDLE <> :HATUAL")
  SQL.Add("  AND GRUPOCONTRATO = :HGRUPOCONTRATO")
  SQL.Add("  AND ((:DIAVENDAINICIAL BETWEEN DIAVENDAINICIAL AND DIAVENDAFINAL) OR")
  SQL.Add("       (:DIAVENDAFINAL BETWEEN DIAVENDAINICIAL And DIAVENDAFINAL) OR")
  SQL.Add("       (DIAVENDAINICIAL BETWEEN :DIAVENDAINICIAL AND :DIAVENDAFINAL))")
  SQL.ParamByName("HATUAL").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("HGRUPOCONTRATO").Value = CurrentQuery.FieldByName("GRUPOCONTRATO").AsInteger
  SQL.ParamByName("DIAVENDAINICIAL").Value = CurrentQuery.FieldByName("DIAVENDAINICIAL").AsInteger
  SQL.ParamByName("DIAVENDAFINAL").Value = CurrentQuery.FieldByName("DIAVENDAFINAL").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    CanContinue = False
    bsShowMessage("Não pode haver intervalos de dias de venda cruzados!", "E")
  End If

  Set SQL = Nothing
End Sub
