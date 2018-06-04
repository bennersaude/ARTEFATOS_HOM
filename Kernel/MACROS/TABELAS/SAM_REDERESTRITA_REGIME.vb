'HASH: 4F4984D0349E8CC0E0BE6F04FCE44493
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim SQL As Object

  Set SQL = NewQuery
  SQL.Active = False
  SQL.Clear
  SQL.Add("SELECT HANDLE FROM SAM_REDERESTRITA_REGIME WHERE REDERESTRITA = :REDERESTRITA AND REGIMEATENDIMENTO = :REGIMEATENDIMENTO")
  SQL.ParamByName("REDERESTRITA").Value = RecordHandleOfTable("SAM_REDERESTRITA")
  SQL.ParamByName("REGIMEATENDIMENTO").Value = CurrentQuery.FieldByName("REDERESTRITA").AsInteger
  SQL.Active = True

  If Not SQL.EOF Then
    bsShowMessage("Regime já cadastrado!", "E")
    CanContinue = False
    Exit Sub
  End If

End Sub
