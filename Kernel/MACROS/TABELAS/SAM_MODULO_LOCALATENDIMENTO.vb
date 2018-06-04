'HASH: 97AB3D8DEAE52CCE338A67B294628F11
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim qVerificaDuplicidade As Object
  Set qVerificaDuplicidade = NewQuery
  qVerificaDuplicidade.Active = False
  qVerificaDuplicidade.Clear
  qVerificaDuplicidade.Add("SELECT Count(1) Encontrou ")
  qVerificaDuplicidade.Add("  FROM SAM_MODULO_LOCALATENDIMENTO ")
  qVerificaDuplicidade.Add(" WHERE HANDLE <> "+CurrentQuery.FieldByName("HANDLE").AsString  )
  qVerificaDuplicidade.Add("   AND LOCALATENDIMENTO =  "+CurrentQuery.FieldByName("LOCALATENDIMENTO").AsString )
  qVerificaDuplicidade.Add("   AND MODULO = "+CurrentQuery.FieldByName("MODULO").AsString )
  qVerificaDuplicidade.Active = True

  If qVerificaDuplicidade.FieldByName("Encontrou").AsInteger > 0 Then
    bsShowMessage("Local de atendimento já cadastrado.", "E")
    CanContinue = False
    Set qVerificaDuplicidade = Nothing
    Exit Sub
  End If

  Set qVerificaDuplicidade = Nothing

End Sub
