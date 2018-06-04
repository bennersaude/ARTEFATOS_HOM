'HASH: 85DBBA42611423F0E6F1CD77E2AB3425
'#Uses "*bsShowMessage"

Dim pPrestador As Long

Sub LookupTipoDoc
	Dim SQL As Object
	Dim VTEXT As String
	Set SQL = NewQuery

	SQL.Add("SELECT TIPOPRESTADOR, CATEGORIA ")
	SQL.Add("  FROM SAM_PRESTADOR ")
	SQL.Add(" WHERE HANDLE = :prestador")

	SQL.ParamByName("prestador").Value = RecordHandleOfTable("SAM_PRESTADOR")
	SQL.Active = True

	If Not SQL.FieldByName("CATEGORIA").IsNull And Not SQL.FieldByName("TIPOPRESTADOR").IsNull Then
		VTEXT = ""
		VTEXT = VTEXT + "   @ALIAS.TIPOPRESTADOR = " + SQL.FieldByName("TIPOPRESTADOR").AsString
		VTEXT = VTEXT + "   AND (EXISTS (SELECT X.HANDLE"
		VTEXT = VTEXT + "                  FROM SAM_TIPOPRESTADOR_DOCCATEGORIA X"
		VTEXT = VTEXT + "                 WHERE X.TIPOPRESTADORDOC = @ALIAS.HANDLE"
		VTEXT = VTEXT + "                   AND X.CATEGORIA = " + SQL.FieldByName("CATEGORIA").AsString
		VTEXT = VTEXT + "                )"
		VTEXT = VTEXT + "        OR NOT EXISTS (SELECT X.HANDLE"
		VTEXT = VTEXT + "                         FROM SAM_TIPOPRESTADOR_DOCCATEGORIA X"
		VTEXT = VTEXT + "                        WHERE X.TIPOPRESTADORDOC = @ALIAS.HANDLE"
		VTEXT = VTEXT + "                      )"
		VTEXT = VTEXT + "       )"
	Else
		If SQL.FieldByName("TIPOPRESTADOR").IsNull Then
			bsShowMessage( "Tipo do prestador não está cadastrado !","I")
			ShowPopup = False
		Else
			bsShowMessage( "Categoria do prestador não está cadastrada !","I")
			ShowPopup = False
		End If
	End If

	UpdateLastUpdate("SAM_TIPOPRESTADOR_DOC")

	If VisibleMode Then
		TIPODOCUMENTO.LocalWhere = Replace(VTEXT, "@ALIAS", "SAM_TIPOPRESTADOR_DOC")
	Else
		TIPODOCUMENTO.WebLocalWhere = VTEXT
	End If
End Sub

Public Sub TABLE_AfterEdit()
	LookupTipoDoc
End Sub

Public Sub TABLE_AfterInsert()
	LookupTipoDoc
End Sub

Public Sub TABLE_AfterPost()
	If Not (CurrentQuery.FieldByName ("DATAENTREGA").IsNull) Then
		WriteAudit("A", HandleOfTable("SAM_PRESTADOR_DOCENTREGUE"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Alterado data de entrega do documento do prestador." + CurrentQuery.FieldByName ("DATAENTREGA").AsString)
	End If

 	Dim callProxy As CSEntityCall
    Set callProxy = BusinessEntity.CreateCall("Benner.Saude.Entidades.Prestadores.SamPrestadorDocEntregue, Benner.Saude.Entidades", "VerificarDocumentosExigidos")

   callProxy.AddParameter(pdtInteger, CurrentQuery.FieldByName("HANDLE").AsInteger)
   callProxy.Execute()

   Set callProxy = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "E", "P", Msg) = "N" Then
		bsShowMessage(Msg, "I")
		CanContinue = False
		Exit Sub
	End If

	If Not (CurrentQuery.FieldByName ("DATAENTREGA").IsNull) Then
		WriteAudit("C", HandleOfTable("SAM_PRESTADOR_DOCENTREGUE"), CurrentQuery.FieldByName("HANDLE").AsInteger, "Excluído documento entregue pelo prestador.")
	End If
End Sub

Public Sub TABLE_AfterScroll()
	If CurrentQuery.FieldByName("TIPODOCUMENTO").AsInteger Then
        Dim SQL As Object
	    Set SQL = NewQuery

		SQL.Add("SELECT CT.MESESVALIDADE ")
		SQL.Add("  FROM SAM_TIPOPRESTADOR_DOC CT ")
		SQL.Add(" WHERE CT.HANDLE = :TIPOPRESTADOR")

		SQL.ParamByName("TIPOPRESTADOR").AsInteger = CurrentQuery.FieldByName("TIPODOCUMENTO").Value

		SQL.Active = True

		If SQL.EOF Then
			bsShowMessage("Não há documento cadastrado para este tipo de prestador.", "E")
			Exit Sub
		End If

		If Not SQL.EOF Then
			If SQL.FieldByName("MESESVALIDADE").AsInteger > 0 Then
				If SQL.FieldByName("MESESVALIDADE").AsInteger = 1 Then
					MESESVALIDADE.Text = "   1 Mês de Validade"
				Else
					MESESVALIDADE.Text = "   " + SQL.FieldByName("MESESVALIDADE").AsString + " Meses de Validade"
				End If
			Else
				MESESVALIDADE.Text = ""
			End If
		End If

		SQL.Active = False

		Set SQL = Nothing
	Else
      MESESVALIDADE.Text = ""
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim qPrest As Object

	Dim SQL As Object
	Set SQL = NewQuery

	If (CurrentQuery.FieldByName("DATAVALIDADE").IsNull) Then

        SQL.Clear
		SQL.Add("SELECT CT.MESESVALIDADE, CT.TABDATA ")
		SQL.Add("  FROM SAM_TIPOPRESTADOR_DOC CT     ")
		SQL.Add(" WHERE CT.HANDLE = :TIPODOCUMENTO   ")

		SQL.ParamByName("TIPODOCUMENTO").AsInteger = CurrentQuery.FieldByName("TIPODOCUMENTO").Value
		SQL.Active = True

		If Not SQL.EOF Then
			If SQL.FieldByName("MESESVALIDADE").IsNull Then
				CurrentQuery.FieldByName("DATAVALIDADE").Clear
			Else
				If SQL.FieldByName("TABDATA").AsString = "D" Then
					CurrentQuery.FieldByName("DATAVALIDADE").Value = DateAdd("m", SQL.FieldByName("MESESVALIDADE").AsInteger, _
					CurrentQuery.FieldByName("DATADOCUMENTO").AsDateTime)
				Else
					If SQL.FieldByName("TABDATA").AsString = "I" Then
						CurrentQuery.FieldByName("DATAVALIDADE").Value = DateAdd("m", SQL.FieldByName("MESESVALIDADE").AsInteger, _
						CurrentQuery.FieldByName("DATAENTREGA").AsDateTime)
					Else
						Set qPrest = NewQuery

						qPrest.Add("SELECT DATACREDENCIAMENTO FROM SAM_PRESTADOR WHERE HANDLE = " + CStr(RecordHandleOfTable("SAM_PRESTADOR")))

						qPrest.Active = True

						If Not qPrest.FieldByName("DATACREDENCIAMENTO").IsNull Then
							CurrentQuery.FieldByName("DATAVALIDADE").Value = DateAdd("m", SQL.FieldByName("MESESVALIDADE").AsInteger, _
							qPrest.FieldByName("DATACREDENCIAMENTO").AsDateTime)
						Else
							CurrentQuery.FieldByName("DATAVALIDADE").Clear
						End If

						Set qPrest = Nothing
					End If
				End If

				If SQL.FieldByName("MESESVALIDADE").AsInteger = 1 Then
					MESESVALIDADE.Text = "   1 Mês de Validade"
				Else
					MESESVALIDADE.Text = "   " + SQL.FieldByName("MESESVALIDADE").AsString + " Meses de Validade"
				End If
			End If
		End If
	End If

    SQL.Clear
    SQL.Add(" SELECT TD.EXIGEANEXO                                               ")
    SQL.Add("   FROM SAM_TIPOPRESTADOR_DOC TPD                                   ")
    SQL.Add("   JOIN SAM_TIPODOCUMENTO     TD  On TD. HANDLE = TPD.TIPODOCUMENTO ")
    SQL.Add("  WHERE TPD.Handle = :TIPODOCUMENTO                                 ")
    SQL.ParamByName("TIPODOCUMENTO").Value = CurrentQuery.FieldByName("TIPODOCUMENTO").AsInteger
    SQL.Active = True

    If ((SQL.FieldByName("EXIGEANEXO").Value = "S") And (CurrentQuery.FieldByName("ARQUIVOANEXO").AsString = "")) Then
	  bsShowMessage("Este tipo de documento exige que seja informado um arquivo anexo.", "E")
      CanContinue = False
      Exit Sub
    End If

	Set SQL = Nothing

	If (Not CurrentQuery.FieldByName("DATAVALIDADE").IsNull) And (CurrentQuery.FieldByName("DATAVALIDADE").AsDateTime <= _
		CurrentQuery.FieldByName("DATAENTREGA").AsDateTime) Then
		bsShowMessage("A data de validade projetada (" + Format(CurrentQuery.FieldByName("DATAVALIDADE").AsDateTime, "dd/mm/yyyy") + _
		              ") deve ser maior que a data de entrega! Verifique as configurações dos documentos exigidos para o Tipo de Prestador.", "E")
		CurrentQuery.FieldByName("DATAVALIDADE").Value = Null
		CanContinue = False
	End If

	If (Not CurrentQuery.FieldByName("DATAVALIDADECERTIDAO").IsNull) And (CurrentQuery.FieldByName("DATAVALIDADECERTIDAO").AsDateTime <= CurrentQuery.FieldByName("DATAEMISSAOCERTIDAO").AsDateTime) Then
		bsShowMessage("A data de validade da certidão deve ser maior que a data de emissão da certidão!", "E")
		CanContinue = False
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

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim Msg As String

	If checkPermissaoFilial (CurrentSystem, "I", "P", Msg) = "N" Then
		bsShowMessage(Msg, "E")
		CanContinue = False
		Exit Sub
	End If
End Sub
