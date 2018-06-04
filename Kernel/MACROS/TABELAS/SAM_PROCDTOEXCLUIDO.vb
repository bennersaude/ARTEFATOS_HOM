'HASH: AABE98934EA479A6025D18EA7F875A00


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.State = 3 Then
    Dim QTemp As Object
    Set QTemp = NewQuery

    QTemp.Active = False
    QTemp.Clear
    QTemp.Add("SELECT COUNT(HANDLE) QT ")
    QTemp.Add("  FROM SAM_PROCDTOEXCLUIDO ")
    QTemp.Add(" WHERE CODIGOANS = :COD")
    QTemp.ParamByName("COD").AsInteger = CurrentQuery.FieldByName("CODIGOANS").AsInteger
    QTemp.Active = True
    If QTemp.FieldByName("QT").AsInteger > 0 Then
      MsgBox("Registro já cadatrado!", vbCritical, "Benner Saúde")
      CanContinue = False
    End If
  End If
End Sub

