'HASH: E606441A62EDFBC4BD5E5E6605AB7E54

Public Sub BOTAOEXCLUIREVENTO_OnClick()
  Dim Obj As Object
  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Excluir(CurrentSystem, "SAM_CID_EVENTO", "Excluindo Eventos para CID", "SAM_TGE", "EVENTO", "CID", CurrentQuery.FieldByName("HANDLE").AsInteger, "", "ESTRUTURA")
  Set Obj = Nothing
  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub


Public Sub BOTAOGERAREVENTOS_OnClick()
  Dim Obj As Object
  Set Obj = CreateBennerObject("SamGerarExcluirDados.Rotinas")
  Obj.Gerar(CurrentSystem, "SAM_CID_EVENTO", "Duplicando Eventos para CID", "SAM_TGE", "EVENTO", "CID", CurrentQuery.FieldByName("HANDLE").AsInteger, "", "ESTRUTURA")
  Set Obj = Nothing
  CurrentQuery.Active = False
  CurrentQuery.Active = True
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim query As Object
	Set query = NewQuery

	query.Add("SELECT COUNT(1) QTDE FROM SAM_CID WHERE ESTRUTURA = :PESTRUTURA AND HANDLE <> :PHANDLE")
	query.ParamByName("PESTRUTURA").AsString = CurrentQuery.FieldByName("ESTRUTURA").AsString
	query.ParamByName("PHANDLE").AsString = CurrentQuery.FieldByName("HANDLE").AsInteger
	query.Active = True

	If query.FieldByName("QTDE").AsInteger > 0 Then
		MsgBox("Já existe um CID com esta estrutura.",vbInformation,"Cid duplicado")
		CanContinue = False
	End If

	Set query = Nothing
End Sub
