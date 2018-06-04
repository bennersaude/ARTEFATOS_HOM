'HASH: 42B7C7421F5A00089257E39C57D26220
'Macro: SAM_PRESTADOR_IRRF
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterInsert()
	TABLE_AfterScroll

	Dim vSelect As String

	vSelect = "(SELECT FISICAJURIDICA FROM SAM_PRESTADOR WHERE HANDLE = "

	If VisibleMode Then
		IRRF.LocalWhere = "FISICAJURIDICA = " + vSelect + "@PRESTADOR)"
		CODIGODIRF.LocalWhere = "FISICAJURIDICA = " + vSelect + "@PRESTADOR)"
	Else
		IRRF.WebLocalWhere = "FISICAJURIDICA = " + vSelect + "@CAMPO(PRESTADOR))"
		CODIGODIRF.WebLocalWhere = "FISICAJURIDICA = " + vSelect + "@CAMPO(PRESTADOR))"
	End If
End Sub

Public Sub TABLE_AfterEdit()
	TABLE_AfterScroll

	Dim vSelect As String

	vSelect = "(SELECT FISICAJURIDICA FROM SAM_PRESTADOR WHERE HANDLE = "

	If VisibleMode Then
		IRRF.LocalWhere = "FISICAJURIDICA = " + vSelect + "@PRESTADOR)"
		CODIGODIRF.LocalWhere = "FISICAJURIDICA = " + vSelect + "@PRESTADOR)"
	Else
		IRRF.WebLocalWhere = "FISICAJURIDICA = " + vSelect + "@CAMPO(PRESTADOR))"
		CODIGODIRF.WebLocalWhere = "FISICAJURIDICA = " + vSelect + "@CAMPO(PRESTADOR))"
	End If
End Sub

Public Sub TABLE_AfterScroll()
	Dim Query As Object
	Set Query = NewQuery

	Query.Clear

	Query.Add("SELECT FISICAJURIDICA FROM SAM_PRESTADOR WHERE HANDLE = :HPRESTADOR")

	Query.ParamByName("HPRESTADOR").AsInteger = CurrentQuery.FieldByName("PRESTADOR").AsInteger
	Query.Active = True

	If Query.FieldByName("FISICAJURIDICA").AsInteger = 1 Then
		TABCONTRIBUICOESFEDERAIS.ReadOnly = True
		NRODEPENDENTES.Visible = True
		OUTRASDEDUCOES.Visible = True
	Else
		TABCONTRIBUICOESFEDERAIS.ReadOnly = False
		NRODEPENDENTES.Visible = False
		OUTRASDEDUCOES.Visible = False
	End If

	Set Query = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim SQLPRE As Object
	Dim SQLIRRF As Object
	Dim SQLDIRF As Object
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If

	Set SQLPRE = NewQuery
	Set SQLIRRF = NewQuery
	Set SQLDIRF = NewQuery

	SQLPRE.Add("SELECT FISICAJURIDICA FROM SAM_PRESTADOR WHERE HANDLE = :PRESTADOR")

	SQLPRE.ParamByName("PRESTADOR").Value = CurrentQuery.FieldByName("PRESTADOR").AsInteger
	SQLPRE.Active = True

	If (Not SQLPRE.EOF) Then
		SQLIRRF.Add("SELECT FISICAJURIDICA FROM SFN_IRRF WHERE HANDLE = :IRRF")

		SQLIRRF.ParamByName("IRRF").Value = CurrentQuery.FieldByName("IRRF").AsInteger
		SQLIRRF.Active = True

		If (Not SQLIRRF.EOF) Then
			If (SQLIRRF.FieldByName("FISICAJURIDICA").AsInteger = 1 And SQLPRE.FieldByName("FISICAJURIDICA").AsInteger <> 1) Then
				Set SQLISS = Nothing
				CanContinue = False
				bsShowMessage("Tipo de IRRF permitido somente para Pessoa Física!", "E")
				Exit Sub
			End If

			If (SQLIRRF.FieldByName("FISICAJURIDICA").AsInteger = 2 And SQLPRE.FieldByName("FISICAJURIDICA").AsInteger <> 2) Then
				Set SQLIRRF = Nothing
				CanContinue = False
				bsShowMessage("Tipo de IRRF permitido somente para Pessoa Jurídica!", "E")
				Exit Sub
			End If
		End If

		SQLDIRF.Add("SELECT FISICAJURIDICA FROM SFN_CODIGODIRF WHERE HANDLE = :DIRF")
		SQLDIRF.ParamByName("DIRF").Value = CurrentQuery.FieldByName("CODIGODIRF").AsInteger
		SQLDIRF.Active = True

		If (Not SQLDIRF.EOF) Then
			If (SQLDIRF.FieldByName("FISICAJURIDICA").AsInteger = 1 And SQLPRE.FieldByName("FISICAJURIDICA").AsInteger <> 1) Then
				Set SQLISS = Nothing
				CanContinue = False
				bsShowMessage("Código DIRF permitido somente para Pessoa Física!", "E")
				Exit Sub
			End If

			If (SQLDIRF.FieldByName("FISICAJURIDICA").AsInteger = 2 And SQLPRE.FieldByName("FISICAJURIDICA").AsInteger <> 2) Then
				Set SQLDIRF = Nothing
				CanContinue = False
				bsShowMessage("Código DIRF permitido somente para Pessoa Jurídica!", "E")
				Exit Sub
			End If
		End If
	End If

	Dim Interface As Object
	Dim Linha As String
	Set Interface = CreateBennerObject("SAMGERAL.Vigencia")

	Linha = Interface.Vigencia(CurrentSystem, "SAM_PRESTADOR_IRRF", "VIGENCIA", "VIGENCIAFINAL", CurrentQuery.FieldByName("VIGENCIA").AsDateTime, CurrentQuery.FieldByName("VIGENCIAFINAL").AsDateTime, "PRESTADOR", "")

	If Linha = "" Then
		CanContinue = True
	Else
		CanContinue = False
		bsShowMessage(Linha, "E")
	End If

	Set Interface = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
