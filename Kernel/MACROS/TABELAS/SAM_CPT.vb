'HASH: 5DA3900EEB48EABF2FA3E1B307BFEFA0
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.State = 3 Then
    Dim QTemp As Object
    Set QTemp = NewQuery

    QTemp.Active = False
    QTemp.Clear
    QTemp.Add("SELECT COUNT(HANDLE) QT ")
    QTemp.Add("  FROM SAM_CPT ")
    QTemp.Add(" WHERE CODIGO = :COD")
    QTemp.ParamByName("COD").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
    QTemp.Active = True
    If QTemp.FieldByName("QT").AsInteger > 0 Then
      bsShowMessage("Registro já cadatrado!", "E")
      CanContinue = False
    End If
  End If
End Sub

