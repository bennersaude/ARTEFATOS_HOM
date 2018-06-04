'HASH: 82A4D517CD455081FE4AA528808BDBC0
'#Uses "*bsShowMessage"

Public Sub ENDERECO_OnPopup(ShowPopup As Boolean)
  ENDERECO.LocalWhere = "PRESTADOR = " + CStr(RecordHandleOfTable("SAM_PRESTADOR"))
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qEnderecoRepetido As Object

  Set qEnderecoRepetido = NewQuery

  qEnderecoRepetido.Active = False
  qEnderecoRepetido.Clear
  qEnderecoRepetido.Add("SELECT HANDLE")
  qEnderecoRepetido.Add("  FROM SAM_MEMBRO_ENDERECO")
  qEnderecoRepetido.Add(" WHERE ENDERECO = :ENDERECO")
  qEnderecoRepetido.Add("   AND HANDLE <> :HANDLE")
  qEnderecoRepetido.Add("   AND CORPOCLINICO = :CORPOCLINICO")
  qEnderecoRepetido.ParamByName("ENDERECO").AsInteger = CurrentQuery.FieldByName("ENDERECO").AsInteger
  qEnderecoRepetido.ParamByName("CORPOCLINICO").AsInteger = CurrentQuery.FieldByName("CORPOCLINICO").AsInteger
  qEnderecoRepetido.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  qEnderecoRepetido.Active = True

  If Not qEnderecoRepetido.EOF Then
    bsShowMessage("Endereço já existente para o membro do corpo clínico!", "A")
    CanContinue = False
  End If

  Set qEnderecoRepetido = Nothing
End Sub
