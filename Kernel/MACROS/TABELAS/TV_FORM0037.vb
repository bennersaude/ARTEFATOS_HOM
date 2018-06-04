'HASH: C8FA06BAE89ACAC9F2D08A7133939D96
'#Uses "*bsShowMessage"

Public Sub TABLE_AfterScroll()
	Dim SQL As Object

	Set SQL = NewQuery
	SQL.Add("SELECT CONTABILIZA FROM SFN_PARAMETROSFIN")
	SQL.Active = True

	If SQL.FieldByName("CONTABILIZA").Value = "S" Then
		ROTCONTABILIZACAO.Visible = True
		DESCHIST.Visible = True
		HISTPADRAO.Visible = True
		DEBITO.Visible = True
		CREDITO.Visible = True
 	Else
 		ROTCONTABILIZACAO.Visible = False
		DESCHIST.Visible = False
		HISTPADRAO.Visible = False
		DEBITO.Visible = False
		CREDITO.Visible = False
 	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)



	Dim SQL As Object
	Dim psDescHist As String
	Dim vsMensagem As String


	Set SQL = NewQuery
	SQL.Add("SELECT HISTORICO FROM SFN_CONTABHIST WHERE HANDLE = :HANDLE")
	SQL.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HISTPADRAO").AsInteger
	SQL.Active = True

	psDescHist = SQL.FieldByName("HISTORICO").AsString

	Dim Interface As Object

	Set Interface = CreateBennerObject("SFNTESOURARIA.Tesouraria")
	vsMensagem = Interface.SaldoInicial(CurrentSystem, CurrentQuery.FieldByName("TESOURARIA").AsInteger, CurrentQuery.FieldByName("DATA").AsDateTime, _
						   CurrentQuery.FieldByName("VALOR").AsFloat, CurrentQuery.FieldByName("HISTORICO").AsString, psDescHist, _
						   CurrentQuery.FieldByName("HISTPADRAO").AsInteger, CurrentQuery.FieldByName("CREDITO").AsInteger)


	Set Interface = Nothing

	If vsMensagem <> "" Then
		bsShowMessage(vsMensagem,"I")
	End If

End Sub
