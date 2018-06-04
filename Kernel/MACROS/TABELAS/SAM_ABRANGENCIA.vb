'HASH: 9303FE64B29F13E578B91E0B952C4A8C
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.State = 3 Then
    Dim SQL As Object
    Set SQL = NewQuery

    SQL.Clear
    SQL.Add("SELECT HANDLE")
    SQL.Add("FROM SAM_ABRANGENCIA")
    SQL.Add("WHERE CODIGO = :CODIGO")
    SQL.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsString
    SQL.Active = True

    If Not SQL.EOF Then
	  bsShowMessage("Já existe este código de Abrangência!","E")
      CanContinue = False
    End If

    Set SQL = Nothing
  End If
End Sub
