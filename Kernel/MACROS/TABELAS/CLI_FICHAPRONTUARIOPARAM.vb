'HASH: 232C011C7B9E00BB37436D481789BF2F


Public Sub TABLE_AfterScroll()
  Dim qFicha As Object
  Set qFicha = NewQuery

  qFicha.Add("SELECT HANDLE FROM CLI_FICHAPRONTUARIO")
  qFicha.Add(" WHERE ARQUIVO = :ARQUIVO")
  qFicha.ParamByName("ARQUIVO").Value = CurrentQuery.FieldByName("ARQUIVO").AsInteger
  qFicha.Active = True

  If qFicha.EOF Then
    FICHA.ReadOnly = False
    GAVETA.ReadOnly = False
  Else
    FICHA.ReadOnly = True
    GAVETA.ReadOnly = True
  End If

  Set qFicha = Nothing
End Sub

