'HASH: FF7FA62CCE76B4805678BC23DE132430
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If CurrentQuery.State = 3 Then
    Dim QTemp As Object
    Set QTemp = NewQuery

    QTemp.Active = False
    QTemp.Clear
    QTemp.Add("SELECT COUNT(HANDLE) QT ")
    QTemp.Add("  FROM SAM_MODULO_PROCDTOEXCLUIDO ")
    QTemp.Add(" WHERE MODULO = :MOD")
    QTemp.Add("   AND PROCDTOEXCLUIDO = :PRO")
    QTemp.ParamByName("MOD").AsInteger = CurrentQuery.FieldByName("MODULO").AsInteger
    QTemp.ParamByName("PRO").AsInteger = CurrentQuery.FieldByName("PROCDTOEXCLUIDO").AsInteger
    QTemp.Active = True
    If QTemp.FieldByName("QT").AsInteger > 0 Then
      bsShowMessage("Procedimento já cadastrado!", "E")
      CanContinue = False
    End If
  End If
End Sub

