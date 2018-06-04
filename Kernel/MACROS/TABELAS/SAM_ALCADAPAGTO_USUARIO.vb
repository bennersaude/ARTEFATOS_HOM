'HASH: 975DB7D975AD0CF3DB0CCD6724F286E0
'#Uses "*bsShowMessage"

'------------------ SMS 90104 - Paulo Melo - 17/12/2007 - INICIO -- nao deve deixar gravar 2 registros do mesmo usuario na mesma alçada
Public Sub TABLE_BeforePost(CanContinue As Boolean)

	Dim USUARIO As Object
	Set USUARIO = NewQuery

	USUARIO.Add("SELECT U.HANDLE")
	USUARIO.Add("FROM SAM_ALCADAPAGTO_USUARIO U")
	USUARIO.Add("JOIN SAM_ALCADAPAGTO A ON (U.ALCADA = A.HANDLE) ")
	USUARIO.Add("WHERE USUARIO = :USUARIO")
	USUARIO.Add("AND ALCADA = :ALCADA")

	USUARIO.ParamByName("USUARIO").AsInteger = CurrentQuery.FieldByName("USUARIO").AsInteger
	USUARIO.ParamByName("ALCADA").AsInteger = CurrentQuery.FieldByName("ALCADA").AsInteger
	USUARIO.Active = True

	If Not USUARIO.EOF Then
		BsShowMessage("Este usuário já está cadastrado nesta alçada.", "E")
		CanContinue = False
	End If

	Set USUARIO = Nothing
'------------------ SMS 90104 - Paulo Melo - 17/12/2007 - FIM
End Sub
