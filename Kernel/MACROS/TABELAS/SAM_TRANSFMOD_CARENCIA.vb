'HASH: 900D0A0515DC91AD16A5F81400DCED8D

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim SQL As Object
  Set SQL = NewQuery

  If CurrentQuery.State = 3 Then

    SQL.Active = False
    SQL.Clear
    SQL.Add("SELECT * FROM SAM_TRANSFMOD_CARENCIA WHERE CARENCIA = :CARENCIA AND TRANSFMOD = :MOD")
    SQL.ParamByName("CARENCIA").Value = CurrentQuery.FieldByName("CARENCIA").AsInteger
    SQL.ParamByName("MOD").Value = CurrentQuery.FieldByName("TRANSFMOD").AsInteger
    SQL.Active = True

    If Not SQL.EOF Then
      MsgBox("Carência já cadastrada.")
      CanContinue = False
      Exit Sub
    End If
  End If

  If CurrentQuery.FieldByName("ACAOTRANSFERENCIA").AsString = "3" Then
    If CurrentQuery.FieldByName("QTDDIASREDEPROPRIA").IsNull Or CurrentQuery.FieldByName("QTDDIASREDECREDENCIADA").IsNull Then
      MsgBox("Quantidade dias rede própria e quantidade dias rede credenciada são obrigatórios quando (Assumir os dias informados.)")
      CanContinue = False
      Exit Sub
    End If

  End If


  Set SQL = Nothing

End Sub


