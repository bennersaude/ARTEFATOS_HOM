'HASH: 14360B585D32DA5073F933F3B97C8EEE

Public Sub TABLE_AfterEdit()
  Dim DePara As Object
  Set DePara = CreateBennerObject("TISSDEPARA.Rotinas")

  DePara.Editar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Set DePara = Nothing

  RefreshNodesWithTable("TIS_DEPARA")
End Sub

