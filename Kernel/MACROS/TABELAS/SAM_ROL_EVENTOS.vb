'HASH: 0030852B1946A9C640E3997E263EB8AA


Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.State = 3 Then
    Dim QTemp As Object
    Set QTemp = NewQuery

    QTemp.Active = False
    QTemp.Clear
    QTemp.Add("SELECT COUNT(HANDLE) QT ")
    QTemp.Add("  FROM SAM_ROL_EVENTOS ")
    QTemp.Add(" WHERE ROL = :ROL")
    QTemp.Add("   AND EVENTO = :EVE")
    QTemp.ParamByName("ROL").AsInteger = CurrentQuery.FieldByName("ROL").AsInteger
    QTemp.ParamByName("EVE").AsInteger = CurrentQuery.FieldByName("EVENTO").AsInteger
    QTemp.Active = True
    If QTemp.FieldByName("QT").AsInteger > 0 Then
      MsgBox("Registro já cadatrado!", vbCritical, "Benner Saúde")
      CanContinue = False
    End If
  End If
End Sub

