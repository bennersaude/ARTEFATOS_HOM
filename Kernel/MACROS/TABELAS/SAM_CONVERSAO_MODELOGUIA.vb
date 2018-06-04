'HASH: C4EF7B9657E36D509EC0A32BEC75F0D8


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim Sql As Object

  Set Sql = NewQuery

  Sql.Add("SELECT * FROM SAM_CONVERSAO_MODELOGUIA WHERE CODIGO = :CODIGO")
  Sql.ParamByName("CODIGO").Value = CurrentQuery.FieldByName("CODIGO").AsString
  Sql.Active = True

  If Not Sql.EOF Then
    MsgBox ("Código já cadastrado")
    CanContinue = False
    Exit Sub
  End If

  Set Sql = Nothing

End Sub

