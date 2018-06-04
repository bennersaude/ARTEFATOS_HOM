'HASH: 7D9C5D6A6D484143D98575C57CB805BA
'#Uses "*bsShowMessage"


Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim qVerificaDuplicidade As Object
  Set qVerificaDuplicidade = NewQuery
  qVerificaDuplicidade.Active = False
  qVerificaDuplicidade.Clear
  qVerificaDuplicidade.Add("SELECT Count(1) Encontrou ")
  qVerificaDuplicidade.Add("  FROM SAM_MODULO_CONDATENDIMENTO ")
  qVerificaDuplicidade.Add(" WHERE HANDLE <> "+CurrentQuery.FieldByName("HANDLE").AsString  )
  qVerificaDuplicidade.Add("   AND CONDICAOATENDIMENTO =  "+CurrentQuery.FieldByName("CONDICAOATENDIMENTO").AsString )
  qVerificaDuplicidade.Add("   AND MODULO = "+CurrentQuery.FieldByName("MODULO").AsString )
  qVerificaDuplicidade.Active = True

  If qVerificaDuplicidade.FieldByName("Encontrou").AsInteger > 0 Then
    bsShowMessage("Condição de atendimento já cadastrado.", "E")
    CanContinue = False
    Set qVerificaDuplicidade = Nothing
    Exit Sub
  End If

  Set qVerificaDuplicidade = Nothing

End Sub
