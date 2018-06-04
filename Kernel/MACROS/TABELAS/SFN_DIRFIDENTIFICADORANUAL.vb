'HASH: F0C2FBB3F7AF60977CB3A0702506E133
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
 Dim sqlConsultaAnoDirf As Object

 Set sqlConsultaAnoDirf = NewQuery

 sqlConsultaAnoDirf.Add(" SELECT ANOBASE                        ")
 sqlConsultaAnoDirf.Add("   FROM SFN_DIRFIDENTIFICADORANUAL     ")
 sqlConsultaAnoDirf.Add("  WHERE ANOBASE = :ANOBASE             ")
 sqlConsultaAnoDirf.Add("    AND HANDLE <> :HANDLE              ")
 sqlConsultaAnoDirf.ParamByName("ANOBASE").AsDateTime = CurrentQuery.FieldByName("ANOBASE").AsDateTime
 sqlConsultaAnoDirf.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
 sqlConsultaAnoDirf.Active = True

 If Not sqlConsultaAnoDirf.EOF Then
 	bsShowMessage("Ano "& Year(CurrentQuery.FieldByName("ANOBASE").AsDateTime) &" já cadastrado!", "E")
 	CanContinue = False
 End If

 Set sqlConsultaAnoDirf = Nothing

End Sub
