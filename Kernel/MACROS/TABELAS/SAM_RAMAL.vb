'HASH: 0C93FCF082E6D2134FDAD400C27CD472
'#Uses "*bsShowMessage"
Option Explicit

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim msg As String
	Dim query As Object
	Set query = NewQuery
	query.Clear
	query.Add("  SELECT RAMAL               ")
	query.Add("    FROM SAM_RAMAL           ")
	query.Add("   WHERE HANDLE <> :HANDLE   ")
	query.Add("     AND RAMAL  =  :RAMAL    ")
	query.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	query.ParamByName("RAMAL").AsString   = CurrentQuery.FieldByName("RAMAL").AsString
	query.Active = True

	If Not query.EOF Then
		msg = "O Ramal : " + CurrentQuery.FieldByName("RAMAL").AsString + " já está cadastrado!"
		bsShowMessage(msg, "E")
		Set query = Nothing
		CanContinue = False
	End If

	Set query = Nothing
End Sub
