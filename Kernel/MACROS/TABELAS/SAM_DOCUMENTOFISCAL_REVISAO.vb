'HASH: 1DBDF9BDB4E2A9D4E85F192DED174AE8
'SAM_DOCUMENTOFISCAL_REVISAO

Option Explicit

'#Uses "*bsShowMessage"


Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	If Not CurrentQuery.FieldByName("ROTINAISSREVISAO").IsNull Then
		CanContinue = False
		bsShowMessage("Registro possui uma rotina de revisão processada. Exclusão não permitida.", "E")
	End If

End Sub

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
	If Not CurrentQuery.FieldByName("ROTINAISSREVISAO").IsNull Then
		CanContinue = False
		bsShowMessage("Registro possui uma rotina de revisão processada. Alteração não permitida.", "E")
	End If
End Sub

Public Sub TABLE_BeforeInsert(CanContinue As Boolean)
	Dim query As BPesquisa
	Set query = NewQuery

	query.Add("SELECT COUNT(1) QTD FROM SAM_DOCUMENTOFISCAL_REVISAO WHERE DOCUMENTOFISCALBASE =:DOCUMENTOFISCALBASE AND HANDLE <> :HANDLE AND ROTINAISSREVISAO IS NULL")
	query.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	query.ParamByName("DOCUMENTOFISCALBASE").AsInteger = RecordHandleOfTable("SAM_DOCUMENTOFISCAL_BASES")
	query.Active = True

	If query.FieldByName("QTD").AsInteger > 0 Then
		CanContinue = False
		bsShowMessage("Existe um registro de revisão ainda não processado para este item.", "E")
	End If



	Set query = Nothing
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If CurrentQuery.FieldByName("ALIQUOTAREVISADA").AsFloat <0 Then
		CanContinue = False
		bsShowMessage("Informe uma alíquota válida para revisão.", "E")
	End If
End Sub
