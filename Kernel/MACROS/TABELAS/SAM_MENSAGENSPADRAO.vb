'HASH: AE655ECD3C21C49E74B2ED93389C858E
 '#uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim msg As String
	msg = ""

	If CurrentQuery.FieldByName("FLAGDOCILEGIVELEMAIL").AsString = "S" Then
		Dim sqle As Object

		Set sqle = NewQuery
		sqle.Add("SELECT ASSUNTO                    ")
		sqle.Add("  FROM SAM_MENSAGENSPADRAO        ")
		sqle.Add(" WHERE FLAGDOCILEGIVELEMAIL ='S'  ")
		sqle.Add("   AND HANDLE <> :HANDLE          ")
		sqle.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger

		sqle.Active = True

		If Not sqle.EOF Then
			msg = msg + "Já existe uma mensagem de E-MAIL padrão para retorno de documento ilegível. Desvincule a mensagem '" _
		        + sqle.FieldByName("ASSUNTO").AsString + "' antes de vincular um novo padrão." + vbCrLf + vbCrLf
		End If

		Set sqle = Nothing
	End If

	If CurrentQuery.FieldByName("FLAGDOCILEGIVELFAX").AsString = "S" Then
		Dim sqlf As Object

		Set sqlf = NewQuery
		sqlf.Add("SELECT ASSUNTO                    ")
		sqlf.Add("  FROM SAM_MENSAGENSPADRAO        ")
		sqlf.Add(" WHERE FLAGDOCILEGIVELFAX ='S'    ")
		sqlf.Add("   AND HANDLE <> :HANDLE          ")
		sqlf.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
		sqlf.Active = True

		If Not sqlf.EOF Then
			msg = msg + "Já existe uma mensagem de FAX padrão para retorno de documento ilegível. Desvincule a mensagem '" _
		        + sqlf.FieldByName("ASSUNTO").AsString + "' antes de vincular um novo padrão."
		End If

		Set sqlf = Nothing
	End If

	If msg <> "" Then
		bsShowMessage(msg, "E")
		CanContinue = False
	End If
End Sub
