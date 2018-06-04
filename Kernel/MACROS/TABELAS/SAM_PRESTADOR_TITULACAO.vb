'HASH: B1D7644AD45AED369CC7C28D04925920
'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim sql As Object
	Set sql = NewQuery

	sql.Clear
	sql.Add("SELECT HANDLE                  ")
	sql.Add("  FROM SAM_PRESTADOR_TITULACAO ")
	sql.Add(" WHERE PRESTADOR = :PRESTADOR  ")
	sql.Add("   AND TITULACAO = :TITULACAO  ")
	sql.Add("   AND HANDLE <> :HANDLE       ")
	sql.ParamByName("PRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
	sql.ParamByName("TITULACAO").AsInteger = CurrentQuery.FieldByName("TITULACAO").AsInteger
	sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	sql.Active = True

	If Not sql.FieldByName("HANDLE").IsNull Then
	  bsShowMessage("Prestador já têm esta Titulação cadastrada!","E")
	  Set sql = Nothing
	  CanContinue = False
	  Exit Sub
	End If

	sql.Clear
	sql.Add("SELECT CODIGO FROM SAM_TITULACAO WHERE HANDLE = :TITULACAO")
	sql.ParamByName("TITULACAO").AsInteger = CurrentQuery.FieldByName("TITULACAO").AsInteger
	sql.Active = True

	Dim vCodigoTitulacao As Integer
	vCodigoTitulacao = sql.FieldByName("CODIGO").AsInteger

	Set sql = Nothing

	If (vCodigoTitulacao > 0) And (CurrentQuery.FieldByName("OBTENCAOAMB").AsString = "N" And CurrentQuery.FieldByName("OBTENCAOCNRM").AsString = "N") Then
	  bsShowMessage("Pelo menos uma das obtenções deve estar marcada!", "E")
	  CanContinue = False
	  Exit Sub
	End If

	If CurrentQuery.FieldByName("OBTENCAOAMB").AsString = "S" And CurrentQuery.FieldByName("DATAOBTENCAOAMB").IsNull Then
	  bsShowMessage("A Data de Obtenção AMB deve ser preenchida quando houve Obtenção da AMB!", "E")
	  CanContinue = False
	  Exit Sub
	End If

	If CurrentQuery.FieldByName("OBTENCAOAMB").AsString = "N" And Not CurrentQuery.FieldByName("DATAOBTENCAOAMB").IsNull Then
	  bsShowMessage("A Data de Obtenção AMB deve estar vazia quando não houve Obtenção da AMB!", "E")
	  CanContinue = False
	  Exit Sub
	End If

	If CurrentQuery.FieldByName("OBTENCAOCNRM").AsString = "S" And CurrentQuery.FieldByName("DATAOBTENCAOCNRM").IsNull Then
	  bsShowMessage("A Data de Obtenção CNRM deve ser preenchida quando houve Obtenção da CNRM!", "E")
	  CanContinue = False
	  Exit Sub
	End If

	If CurrentQuery.FieldByName("OBTENCAOCNRM").AsString = "N" And Not CurrentQuery.FieldByName("DATAOBTENCAOCNRM").IsNull Then
	  bsShowMessage("A Data de Obtenção CNRM deve estar vazia quando não houve Obtenção da CNRM!", "E")
	  CanContinue = False
	  Exit Sub
	End If

End Sub
