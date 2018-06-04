'HASH: 80BC430E124320A032B23C687D828556
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.State = 3 Then
    Dim SQL As Object
    Set SQL = NewQuery

    SQL.Clear
    SQL.Add("SELECT HANDLE")
    SQL.Add("FROM SAM_SEGMENTACAO")
    SQL.Add("WHERE CODIGO = :CODIGO")
    SQL.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsString
    SQL.Active = True

    If Not SQL.EOF Then
      bsShowMessage("Já existe este Código de Segmentação cadastrado!","E")
      CanContinue = False
    End If

    Set SQL = Nothing
  End If
End Sub
