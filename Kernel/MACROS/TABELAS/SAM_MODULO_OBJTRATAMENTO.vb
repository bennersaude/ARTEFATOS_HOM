'HASH: 2386546622FF05E998EA4D4DE2B4622A
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)

  Dim qVerificaDuplicidade As Object
  Set qVerificaDuplicidade = NewQuery
  qVerificaDuplicidade.Active = False
  qVerificaDuplicidade.Clear
  qVerificaDuplicidade.Add("SELECT Count(1) Encontrou ")
  qVerificaDuplicidade.Add("  FROM SAM_MODULO_OBJTRATAMENTO ")
  qVerificaDuplicidade.Add(" WHERE HANDLE <> "+CurrentQuery.FieldByName("HANDLE").AsString  )
  qVerificaDuplicidade.Add("   AND OBJETIVOTRATAMENTO =  "+CurrentQuery.FieldByName("OBJETIVOTRATAMENTO").AsString )
  qVerificaDuplicidade.Add("   AND MODULO = "+CurrentQuery.FieldByName("MODULO").AsString )
  qVerificaDuplicidade.Active = True

  If qVerificaDuplicidade.FieldByName("Encontrou").AsInteger > 0 Then
    bsShowMessage("Objetivo de atendimento já cadastrado.", "E")
    CanContinue = False
    Set qVerificaDuplicidade = Nothing
    Exit Sub
  End If

  Set qVerificaDuplicidade = Nothing

End Sub
