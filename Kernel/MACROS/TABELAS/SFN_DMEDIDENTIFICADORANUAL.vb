'HASH: 6F62464128F97758B0E3D941AC281C12
 

'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
 Dim CONSULTA As Object

 Set CONSULTA = NewQuery

 CONSULTA.Add(" SELECT ANOCALENDARIO                  ")
 CONSULTA.Add("   FROM SFN_DMEDIDENTIFICADORANUAL     ")
 CONSULTA.Add("  WHERE ANOCALENDARIO = :ANOCALENDARIO ")
 CONSULTA.Add("    AND HANDLE <> :HANDLE              ")
 CONSULTA.ParamByName("ANOCALENDARIO").AsDateTime = CurrentQuery.FieldByName("ANOCALENDARIO").AsDateTime
 CONSULTA.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
 CONSULTA.Active = True

 If Not CONSULTA.EOF Then
 	bsShowMessage("Ano "& Year(CurrentQuery.FieldByName("ANOCALENDARIO").AsDateTime) &" já cadastrado!", "E")
 	CanContinue = False
 End If

 Set CONSULTA = Nothing
End Sub
