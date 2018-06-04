'HASH: BC76583FCF72EDE970DC1DA008797586
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  'Balani SMS 54421 09/01/2006

  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    If CurrentQuery.FieldByName("DATAFINAL").AsDateTime < CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
      bsShowMessage("Data final menor que a data inicial.", "E")
      CanContinue = False
      Exit Sub
    End If
  End If

  Dim SQL As Object
  Set SQL = NewQuery

  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE")
  SQL.Add("  FROM SAM_CONTRATO_AUXILIOPACOTE")
  SQL.Add(" WHERE CONTRATOAUXILIO = :CONTRATOAUXILIO")
  SQL.Add("   AND PACOTEAUXILIO = :PACOTEAUXILIO")
  SQL.Add("   AND HANDLE <> :HANDLE")
  SQL.Add("   AND DATAINICIAL <= :DATAINICIAL")
  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    SQL.Add("   AND (DATAFINAL IS NULL OR DATAFINAL >= :DATAFINAL)")
  End If
  SQL.ParamByName("CONTRATOAUXILIO").AsInteger = CurrentQuery.FieldByName("CONTRATOAUXILIO").AsInteger
  SQL.ParamByName("PACOTEAUXILIO").AsInteger = CurrentQuery.FieldByName("PACOTEAUXILIO").AsInteger
  SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  SQL.ParamByName("DATAINICIAL").AsDateTime = CurrentQuery.FieldByName("DATAINICIAL").AsDateTime
  If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
    SQL.ParamByName("DATAFINAL").AsDateTime = CurrentQuery.FieldByName("DATAFINAL").AsDateTime
  End If
  SQL.Active = True

  If Not SQL.FieldByName("HANDLE").IsNull Then
    CanContinue = False
    bsShowMessage("Existe outra vigência com data final aberta.", "E")
  End If

  Set SQL = Nothing
  'final SMS 54421
End Sub
