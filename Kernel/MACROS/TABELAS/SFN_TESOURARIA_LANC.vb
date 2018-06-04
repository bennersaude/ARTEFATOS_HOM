'HASH: F2C190C47B67ECAF61417FC819647068
'#Uses "*bsShowMessage"


Public Sub BOTAOCONFERENCIA_OnClick()
  Dim Interface As Object

  Set Interface = CreateBennerObject("SFNTesouraria.Rotinas")
  Interface.Chama(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "Conferido")
  Set Interface = Nothing

End Sub

Public Sub BOTAOCONSULTAR_OnClick()
  Dim Interface As Object

  Set Interface = CreateBennerObject("SFNTesouraria.Rotinas")
  Interface.Chama(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "Consulta")
  Set Interface = Nothing

End Sub

Public Sub BOTAOESTORNAR_OnClick()
  Dim Interface As Object

  If Not CurrentQuery.FieldByName("CANCELADODATA").IsNull Then
    bsShowMessage("Lançamento está Cancelado!","I")
    Exit Sub
  End If

  If Not CurrentQuery.FieldByName("CONFERIDODATA").IsNull Then
    bsShowMessage("Lançamento já foi Conferido!", "I")
    Exit Sub
  End If

  Set Interface = CreateBennerObject("SFNTesouraria.Rotinas")
  Interface.Chama(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger, "Estorno")
  Set Interface = Nothing
  RefreshNodesWithTable("SFN_TESOURARIA_LANC")
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
  Dim Sql1 As Object
  Set Sql1 = NewQuery

  Dim Sql As Object
  Set Sql = NewQuery

  Sql1.Add("SELECT HANDLE FROM SFN_DOCUMENTO")
  Sql1.Add("WHERE TESOURARIALANC = :PLANCTES")
  Sql1.ParamByName("PLANCTES").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Sql1.Active = True
  If Not Sql.EOF Then
    bsShowMessage("Só podem ser excluídos lançamentos de Transferência e/ou Saldo inicial", "E")
    CanContinue = False
    Set Sql = Nothing
    Exit Sub
  End If

  Sql.Add("SELECT HANDLE FROM SFN_FATURA_LANC")
  Sql.Add("WHERE TESOURARIALANC = :PLANCTES")
  Sql.ParamByName("PLANCTES").Value = CurrentQuery.FieldByName("HANDLE").AsInteger
  Sql.Active = True

  If Not Sql.EOF Then
    bsShowMessage("Só podem ser excluídos lançamentos de Transferência e/ou Saldo inicial", "E")
    CanContinue = False
    Set Sql1 = Nothing
    Exit Sub
  End If

	Dim qDoc As Object
	Dim qFat As Object

	Set qDoc = NewQuery
	Set qFat = NewQuery

	qDoc.Add("SELECT * FROM SFN_DOCUMENTO")
	qDoc.Add("WHERE TESOURARIALANC = :PDOC")
	qDoc.ParamByName("PDOC").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qDoc.Active = True

	qFat.Add("SELECT DISTINCT F.HANDLE,          	")
	qFat.Add("                F.DATAVENCIMENTO,     ")
	qFat.Add("				  F.DATAEMISSAO,        ")
	qFat.Add("				  F.DATACONTABIL,       ")
	qFat.Add("				  F.BAIXADATA,          ")
	qFat.Add("				  F.NUMERO,          	")
	qFat.Add("				  F.VALOR,         		")
	qFat.Add("				  F.NATUREZA,           ")
	qFat.Add("				  F.BAIXAJURO,          ")
	qFat.Add("				  F.BAIXAMULTA,         ")
	qFat.Add("                F.BAIXACORRECAO       ")
	qFat.Add("   FROM  SFN_FATURA_LANC FL,          ")
	qFat.Add("	       SFN_FATURA F         		")
	qFat.Add(" WHERE FL.TESOURARIALANC = :PFAT AND  ")
	qFat.Add("		 F.HANDLE = FL.FATURA           ")
	qFat.ParamByName("PFAT").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qFat.Active = True


	Dim qContab As Object
	Set qContab = NewQuery

	qContab.Add("SELECT HANDLE FROM SFN_CONTAB_LANC")
	qContab.Add("WHERE TESOURARIAMOV = :PHANDLE")

	Dim qExcluirDebCre As Object
	Set qExcluirDebCre = NewQuery

	qExcluirDebCre.Add("DELETE FROM SFN_CONTAB_LANC_DEBCRE")
	qExcluirDebCre.Add(" WHERE CONTABLANC = :PCONTAB")

	Dim qExcluirContab As Object
	Set qExcluirContab = NewQuery

	qExcluirContab.Add("DELETE FROM SFN_CONTAB_LANC")
	qExcluirContab.Add(" WHERE TESOURARIAMOV = :pHANDLE")

	qContab.ParamByName("PHANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	qContab.Active = True


	If (qDoc.EOF) And (qFat.EOF) Then
	  		qExcluirDebCre.ParamByName("PCONTAB").AsInteger = qContab.FieldByName("HANDLE").AsInteger
	    	qExcluirDebCre.ExecSQL
			qExcluirContab.ParamByName("Phandle").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
			qExcluirContab.ExecSQL
	Else
			Dim qCampos As Object
			Set qCampos = NewQuery

			qCampos.Add("SELECT * FROM SFN_TESOURARIA_LANC")
			qCampos.Add("WHERE HANDLE = :PLANC			  ")
			qCampos.ParamByName("PLANC").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger


			qExcluirDebCre.ParamByName("PCONTAB").AsInteger = qContab.FieldByName("HANDLE").AsInteger
	    	qExcluirDebCre.ExecSQL
			qExcluirContab.ParamByName("Phandle").AsInteger = qCampos.FieldByName("TRANSFTESOURARIALANC").AsInteger
			qExcluirContab.ExecSQL
	End If



End Sub

Public Sub TABLE_OnCommandClick(ByVal CommandID As String, CanContinue As Boolean)
	Select Case CommandID
		Case "BOTAOCONFERENCIA"
			BOTAOCONFERENCIA_OnClick
		Case "BOTAOCONSULTAR"
			BOTAOCONSULTAR_OnClick
	End Select
End Sub
