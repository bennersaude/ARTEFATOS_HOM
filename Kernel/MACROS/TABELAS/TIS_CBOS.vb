'HASH: 465215958D19242CB9B4025697ED0173
Option Explicit

'#Uses "*bsShowMessage"
'#Uses "*RegistrarLogAlteracao"

Dim especialidadeBk As Long

Public Sub TABLE_AfterScroll()
  especialidadeBk = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If Not CurrentQuery.FieldByName("ESPECIALIDADE").IsNull Then
		Dim qAux As Object
		Set qAux = NewQuery

		qAux.Add("SELECT CODIGO, DESCRICAO")
		qAux.Add("  FROM TIS_CBOS")
		qAux.Add(" WHERE ESPECIALIDADE = :ESPECIALIDADE")
		qAux.Add("   AND VERSAOTISS = :VERSAOTISS")

		qAux.ParamByName("ESPECIALIDADE").AsInteger = CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger
		qAux.ParamByName("VERSAOTISS").AsInteger = CurrentQuery.FieldByName("VERSAOTISS").AsInteger

		qAux.Active = True

		If Not qAux.EOF Then
			bsShowMessage("O CBO-S '" + _
				qAux.FieldByName("CODIGO").AsString + " - " + _
				qAux.FieldByName("DESCRICAO").AsString + "' já possui a especialidade selecionada!", "E")

			CanContinue = False

			Exit Sub
		End If
	End If
End Sub

Public Sub TABLE_AfterPost()
    If CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger <> especialidadeBk Then
      If especialidadeBk <> 0 Then
        RegistrarLogAlteracao "SAM_ESPECIALIDADE", especialidadeBk, "TIS_CBOS TABLE_AfterPost"
      End If
      If CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger <> 0 Then
        RegistrarLogAlteracao "SAM_ESPECIALIDADE", CurrentQuery.FieldByName("ESPECIALIDADE").AsInteger, "TIS_CBOS TABLE_AfterPost"
      End If
    End If

End Sub
