'HASH: 8510F0D4D6F5B7EFE1E6496B48D37876
'ATENÇAO: HÁ IMPLEMENTAÇÕES BEF PARA ESTA MACRO

'#Uses "*bsShowMessage"

Option Explicit

Public Sub TABLE_AfterScroll()
	DENTESPERMITIDOS.Visible = CurrentQuery.FieldByName("TIPO").AsString = "R"

	AtualizouVerif12(False)
End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	CanContinue = AtualizouVerif12(False)
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	CanContinue = AtualizouVerif12(True)
End Sub

Public Sub TABLE_NewRecord()
	DENTESPERMITIDOS.Visible = False
	CurrentQuery.FieldByName("TIPO").AsString = "D"
End Sub

Public Sub TIPO_OnChange()
	CurrentQuery.UpdateRecord
	DENTESPERMITIDOS.Visible = CurrentQuery.FieldByName("TIPO").AsString = "R"
End Sub

Public Function AtualizouVerif12(EhInsercao As Boolean) As Boolean
	AtualizouVerif12 = True

	If EhInsercao Then
		Dim handlesVerif12 As String
		handlesVerif12 = ",1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,116,117,118,119,120,200,201,202,203,204,205,206,207,300,301,302,303,304,305,306,307,308,309,310,311,400,401,402,403,404,405,406,407,408,409,410,411,412,413,414,415,"
    	'evitar conflito de handles para não prejudicar a BSCLI006.TCadastrarGrauFrm.GerarGraus ou o Odontograma do PEP

		If InStr(1, handlesVerif12, ","+Trim(CurrentQuery.FieldByName("HANDLE").AsString)+",") > 0 Then
	    	bsShowMessage("Tabela de dentes está desatualizada. Processar a verificação '12 - Verificação para atualizar a tabela de sistema com o cadastro de dentes' para visualização correta dos dados odontológicos!", "E")
	    	AtualizouVerif12 = False
		End If
	End If

	If Len(Trim(CurrentQuery.FieldByName("TIPO").AsString)) = 0 Then
	    bsShowMessage("Tabela de dentes está desatualizada. Processar a verificação '12 - Verificação para atualizar a tabela de sistema com o cadastro de dentes' para visualização correta dos dados odontológicos!", "E")
	    AtualizouVerif12 = False
	End If
End Function
