'HASH: 892FAE10B7DF1C0F753CD68138CA1043
 

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  If VisibleMode Or WebVisionCode <> "" Then
    Dim sqlRecuperaRegistro As Object

    Set sqlRecuperaRegistro = NewQuery
    sqlRecuperaRegistro.Clear
    sqlRecuperaRegistro.Active = False
    sqlRecuperaRegistro.Add("SELECT LATITUDE, LONGITUDE FROM SAM_PRESTADOR_ENDERECO WHERE HANDLE = :HANDLE")
    sqlRecuperaRegistro.ParamByName("HANDLE").Value = RecordHandleOfTable("SAM_PRESTADOR_ENDERECO")
    sqlRecuperaRegistro.Active = True

    If Not(CurrentQuery.FieldByName("LATITUDE").IsNull) Or Not(CurrentQuery.FieldByName("LONGITUDE").IsNull) Then
      If CurrentQuery.FieldByName("LATITUDE").AsFloat <> sqlRecuperaRegistro.FieldByName("LATITUDE").AsFloat Or CurrentQuery.FieldByName("LONGITUDE").AsFloat <> sqlRecuperaRegistro.FieldByName("LONGITUDE").AsFloat Then
        CurrentQuery.FieldByName("DTATUALIZACAOLATITUDELONGITUDE").AsDateTime = CurrentSystem.ServerDate
      End If
      If CurrentQuery.FieldByName("LATITUDE").AsFloat <> sqlRecuperaRegistro.FieldByName("LATITUDE").AsFloat Or CurrentQuery.FieldByName("LONGITUDE").AsFloat <> sqlRecuperaRegistro.FieldByName("LONGITUDE").AsFloat Then
        CurrentQuery.FieldByName("DTATUALIZACAOLATITUDELONGITUDE").AsDateTime = CurrentSystem.ServerDate
      End If
    End If
    Set sqlRecuperaRegistro = Nothing
  End If
End Sub
