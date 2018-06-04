'HASH: 461E2C1D096B657D3B1EC25229EE068C


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim QUpdate As Object

  Set QUpdate = NewQuery
  QUpdate.Clear
  QUpdate.Add("UPDATE AEX_PRESTADOR_PRS SET")
  QUpdate.Add("       PROCESSADO = 'N'")
  QUpdate.Add(" WHERE EMPCONECT = :EMPCO")
  QUpdate.Add("   AND CODPRESTADOR = :CODPRESTADOR")
  QUpdate.Add("   AND HANDLE = :HNDL")
  QUpdate.ParamByName("EMPCO").Value = CurrentQuery.FieldByName("EMPCONECT").AsInteger
  QUpdate.ParamByName("CODPRESTADOR").Value = CurrentQuery.FieldByName("CODPRESTADOR").AsString
  QUpdate.ParamByName("HNDL").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  QUpdate.ExecSQL
  Set QUpdate = Nothing
End Sub

