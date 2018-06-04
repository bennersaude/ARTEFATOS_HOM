'HASH: A0085682E251FAB8AE5A061C11C425CD


Public Sub TABLE_AfterScroll()
  Dim qFicha As Object
  Set qFicha = NewQuery

  qFicha.Add("SELECT HANDLE FROM CLI_FICHAPRONTUARIO")
  qFicha.Add(" WHERE ARQUIVO = :ARQUIVO")
  qFicha.ParamByName("ARQUIVO").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  qFicha.Active = True

  If qFicha.EOF Then
    CODIGO.ReadOnly = False
  Else
    CODIGO.ReadOnly = True
  End If

  Set qFicha = Nothing
End Sub

