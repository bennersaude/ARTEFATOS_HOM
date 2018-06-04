'HASH: 9C74073662B5A0290AD2A6060D00AC06

'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim Query As Object
	Set Query = NewQuery
	Query.Add("SELECT DEsCRICAO			  ")
	Query.Add("  FROM SAM_TABELASITUACAORH")
	Query.Add(" WHERE CODIGO = :CODIGORH")
	Query.Add("   AND HANDLE <> :HANDLE")
	Query.ParamByName("CODIGORH").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
	Query.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
	Query.Active = True
	If Not (Query.FieldByName("DESCRICAO").IsNull) Then
    	bsShowMessage("Código Informado Já Existe", "E")
    	CanContinue = False
    End If
End Sub
