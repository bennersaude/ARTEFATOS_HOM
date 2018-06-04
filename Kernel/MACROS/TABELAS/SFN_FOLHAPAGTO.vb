'HASH: 8E87BA92A55F386DF39EBEFEBC07816D

'#Uses "*bsShowMessage"

Public Sub TABLE_BeforePost(CanContinue As Boolean)
 '------------------- SMS 90104 - Paulo Melo - 17/12/2007 - INICIO  -- Fazer o tratamento para nao dar erro de unique do banco, mas sim mostra uma mensagem.
 Dim CODIGO As Object
 Set CODIGO = NewQuery

 CODIGO.Add("SELECT HANDLE")
 CODIGO.Add("FROM SFN_FOLHAPAGTO")
 CODIGO.Add("WHERE CODIGO = :CODIGO")
 CODIGO.Add("AND HANDLE <> :HANDLE")	' SMS 87427 - Paulo Melo - ao editar e salvar, bloqueava o proprio registro.
 CODIGO.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
 CODIGO.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
 CODIGO.Active = True

 If Not CODIGO.EOF Then
 	bsShowMessage("Já existe uma folha com esse código", "E")
 	CanContinue = False
 End If

 Set CODIGO = Nothing
'------------------- SMS 90104 - Paulo Melo - 17/12/2007 - FIM
End Sub
 
