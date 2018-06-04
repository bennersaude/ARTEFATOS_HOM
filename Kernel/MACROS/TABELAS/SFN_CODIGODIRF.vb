'HASH: 4B9EFB54955C510863FB2053FB22EE18
'#Uses "*bsShowMessage"


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  If CurrentQuery.State <> 2 Then
    SQL.Clear
    SQL.Active = False
    SQL.Add("SELECT HANDLE FROM SFN_CODIGODIRF WHERE CODIGORETENCAO = :CODIGO")
    SQL.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGORETENCAO").AsInteger
    SQL.Active = True

    If Not SQL.EOF Then
      bsShowMessage("Já existe um registro com este mesmo código DIRF !", "E")
      CanContinue = False
    End If
  End If

  If CurrentQuery.State = 2 Then
    SQL.Clear
    SQL.Active = False
    SQL.Add("SELECT HANDLE FROM SFN_CODIGODIRF WHERE CODIGORETENCAO = :CODIGO")
    SQL.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGORETENCAO").AsInteger
    SQL.Active = True

    If (CurrentQuery.FieldByName("HANDLE").AsInteger <> SQL.FieldByName("HANDLE").AsInteger) And (Not SQL.EOF) Then
      bsShowMessage("Já existe um registro com este mesmo código DIRF !", "E")
      CanContinue = False
    End If
  End If

  Set SQL = Nothing

End Sub

