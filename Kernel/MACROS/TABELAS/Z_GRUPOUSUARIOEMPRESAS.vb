'HASH: C5C371B650606F824656977B4B4047A3
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
	If WebMode Then
		If WebVisionCode = "V_Z_GRUPOUSUARIOEMPRESAS_366" Then
			USUARIO.ReadOnly = True
			GRUPO.ReadOnly = True
		End If
	End If
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Q As Object
  Set Q = NewQuery
  Q.Add("SELECT COUNT(*) N FROM Z_GRUPOUSUARIOEMPRESAFILIAIS WHERE USUARIO = :USUARIO AND EMPRESA = :EMPRESA")
  Q.ParamByName("USUARIO").Value = CurrentQuery.FieldByName("USUARIO").Value
  Q.ParamByName("EMPRESA").Value = CurrentQuery.FieldByName("EMPRESA").Value
  Q.Active = True
  If Q.FieldByName("N").AsInteger > 0 Then
    CanContinue = False
    bsShowMessage("Registro não pode ser excluído porque Usuário possui direitos para filiais dessa empresa.", "E")
  End If
  Q.Active = False
  Set Q = Nothing
End Sub

