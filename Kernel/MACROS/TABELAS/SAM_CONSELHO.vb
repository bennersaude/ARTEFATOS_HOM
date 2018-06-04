'HASH: E662F2C87FCC3B2052AB09D016A3D4C1

'#Uses "*bsShowMessage"
Public Sub TABLE_BeforePost(CanContinue As Boolean)
  Dim qVerifica As BPesquisa
  Set qVerifica = NewQuery

  If Not CurrentQuery.FieldByName("CODIGOTISS").IsNull And CurrentQuery.FieldByName("CODIGOTISS").AsString <> "10" Then
	qVerifica.Clear
	qVerifica.Add("SELECT COUNT(1) QTD              ")
	qVerifica.Add("  FROM SAM_CONSELHO              ")
	qVerifica.Add(" WHERE CODIGOTISS = :CODIGOTISS  ")
	qVerifica.Add("   AND HANDLE <> :HANDLE         ")
	qVerifica.ParamByName("CODIGOTISS").AsString = CurrentQuery.FieldByName("CODIGOTISS").AsString
	qVerifica.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qVerifica.Active = True

	If qVerifica.FieldByName("QTD").AsInteger > 0 Then
	  BsShowMessage("Já existe um conselho com este código TISS!", "I")
	  CanContinue = False
	End If
  End If

  Set qVerifica = Nothing
End Sub
