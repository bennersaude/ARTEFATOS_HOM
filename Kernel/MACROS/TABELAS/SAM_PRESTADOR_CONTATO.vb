'HASH: C5C442F29CF633343352B0A3B41537BE
'Macro: SAM_PRESTADOR_CONTATO
'#Uses "*bsShowMessage"

Public Sub CEP_OnPopup(ShowPopup As Boolean)
	' Joldemar Moreira 12/06/2003
	' SMS 16059
	Dim vHandle As String
	Dim interface As Object

	ShowPopup = False

	Set interface = CreateBennerObject("ProcuraCEP.Rotinas")

	interface.Exec(CurrentSystem, vHandle)

	If vHandle <>"" Then
		Dim SQL As Object
		Set SQL = NewQuery

		SQL.Add("SELECT CEP,ESTADO,MUNICIPIO,BAIRRO,LOGRADOURO,COMPLEMENTO   ")
		SQL.Add("  FROM LOGRADOUROS      ")
		SQL.Add(" WHERE CEP = :HANDLE ")

		SQL.ParamByName("HANDLE").Value = vHandle
		SQL.Active = True

		CurrentQuery.Edit
		CurrentQuery.FieldByName("CEP").Value = SQL.FieldByName("CEP").AsString
		CurrentQuery.FieldByName("ESTADO").Value = SQL.FieldByName("ESTADO").AsString
		CurrentQuery.FieldByName("MUNICIPIO").Value = SQL.FieldByName("MUNICIPIO").AsString
		CurrentQuery.FieldByName("BAIRRO").Value = SQL.FieldByName("BAIRRO").AsString
		CurrentQuery.FieldByName("LOGRADOURO").Value = SQL.FieldByName("LOGRADOURO").AsString
		CurrentQuery.FieldByName("COMPLEMENTO").Value = SQL.FieldByName("COMPLEMENTO").AsString
	End If

	Set interface = Nothing
End Sub

Public Sub TABLE_AfterScroll()

	TIPO.ReadOnly = False

End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim SQL As Object
	Set SQL = NewQuery

	SQL.Add("SELECT EXIGECPFCNPJ")
	SQL.Add("  FROM	SAM_RELACAOEMPRESARIAL")
	SQL.Add(" WHERE HANDLE = :HANDLEREL")

	SQL.ParamByName("HANDLEREL").Value = CurrentQuery.FieldByName("RELACAOEMPRESARIAL").AsInteger
	SQL.Active = True

	If Not(SQL.EOF)Then
		If SQL.FieldByName("EXIGECPFCNPJ").Value = "S" Then
			If CurrentQuery.FieldByName("CPF").IsNull Then
				bsShowMessage("Relação Empresarial exige preenchimento de CPF", "E")

				CanContinue = False
				Exit Sub
			End If
		End If
	End If

	Set SQL = Nothing

	If Not(CurrentQuery.FieldByName("CPF").IsNull)Then
		If Len(CurrentQuery.FieldByName("CPF").AsString) = 11 Then
			If Not IsValidCPF(CurrentQuery.FieldByName("CPF").AsString)Then
				bsShowMessage("CPF Inválido", "E")
				CanContinue = False
				Exit Sub
			End If
		Else
			bsShowMessage("CPF Inválido", "E")
			CanContinue = False
			Exit Sub
		End If
	End If

	If Not CurrentQuery.FieldByName("DATAFINAL").IsNull Then
		If CurrentQuery.FieldByName("DATAFINAL").AsDateTime <CurrentQuery.FieldByName("DATAINICIAL").AsDateTime Then
			bsShowMessage("Data Final nao pode ser menor que a data Inicial", "E")
			CanContinue = False
		End If
	End If

	If TIPO.PageIndex = 2 Then
	 If (CurrentQuery.FieldByName("EMAIL").AsString = "") Or (CurrentQuery.FieldByName("TELEFONE").AsString = "") Then
		   	bsShowMessage("Quando o tipo selecionado for Responsável Técnico o campo e-mail e telefone devem ser preenchidos", "E")
		   	CanContinue = False
	 End If
	End If

	If WebMode Then
	  If CurrentQuery.FieldByName("TIPO").AsString = "3" Then
	  	If (CurrentQuery.FieldByName("EMAIL").AsString = "") Or (CurrentQuery.FieldByName("TELEFONE").AsString = "") Then
		   	bsShowMessage("Quando o tipo selecionado for Responsável Técnico o campo e-mail e telefone devem ser preenchidos", "E")
		   	CanContinue = False
	 	End If
	  End If
	End If

End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "A", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial(CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub

Public Sub TABLE_NewRecord()
	If WebMode Then
		CurrentQuery.FieldByName("PRESTADOR").AsInteger = RecordHandleOfTable("SAM_PRESTADOR")
	End If
End Sub

Public Sub TIPO_OnChange()
  If TIPO.PageIndex = 2 Then
    TELEFONE.Hint = "Telefone do responsável técnico do prestador."
    EMAIL.Hint = "E-mail do responsável técnico do prestador."
  End If

  If TIPO.PageIndex = 1 Then
    TELEFONE.Hint = "Telefone do representante legal do prestador."
    EMAIL.Hint = "E-mail do representante legal do prestador."
  End If

  If TIPO.PageIndex = 0 Then
    TELEFONE.Hint = "Telefone do contato."
    EMAIL.Hint = "E-mail do contato do prestador."
  End If

End Sub
