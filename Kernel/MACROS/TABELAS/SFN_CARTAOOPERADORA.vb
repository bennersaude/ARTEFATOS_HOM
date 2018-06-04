'HASH: 5298E1A3AAA2F994D9C4B9F31208092C
'#Uses "*bsShowMessage"
'SFN_CARTAOOPERADORA

Dim vCodigoCartao As Integer

Public Sub TABLE_BeforeEdit(CanContinue As Boolean)
 vCodigoCartao = CurrentQuery.FieldByName("CODIGO").AsInteger
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)

Dim qCartao As Object
Set qCartao = NewQuery

qCartao.Add("SELECT HANDLE")
qCartao.Add("FROM SFN_CARTAOOPERADORA")
qCartao.Add("WHERE CODIGO = :CODIGO")
qCartao.ParamByName("CODIGO").AsInteger = CurrentQuery.FieldByName("CODIGO").AsInteger
qCartao.Active = True

If CurrentQuery.State = 3 And Not qCartao.EOF Then
	bsShowMessage("Já existe um cartão com esse código", "E")
	CanContinue = False
End If

If CurrentQuery.State = 2 And vCodigoCartao <> CurrentQuery.FieldByName("CODIGO").AsInteger And qCartao.FieldByName("HANDLE").AsInteger > 0 Then
	bsShowMessage("Já existe um cartão com esse código", "E")
	CanContinue = False
End If

Set qCartao = Nothing

End Sub
 
