'HASH: F79978A9E958D0CCDC7E22995B6189C4
'TABELA: TIS_MENSAGEMTISS

Option Explicit


Public Sub BOTAOPROCESSAR_OnClick()
	Dim obj As Object

	If CurrentQuery.State = 3 Then
		MsgBox("Os dados não podem estar em edição!")
		Exit Sub
	End If

 	If CurrentQuery.FieldByName("SITUACAO").AsString = "A" Then 'somente situação aberta

		Set obj = CreateBennerObject("BSTISS.Rotinas")
		obj.ProcessarRotina(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger) 'Irá processar o ARQUIVO de ARQUIVO com o tipo de transação.
		Set obj = Nothing


	Else
		If VisibleMode Then
			MsgBox("Somente pode ser processado na situação AGUARDANDO")
		Else
			CancelDescription = "Somente pode ser processado na situação AGUARDANDO"
		End If
	End If

	CurrentQuery.Active = False
	CurrentQuery.Active = True
End Sub


Public Sub TABLE_AfterCommitted()
  Dim obj As Object
  Set obj=CreateBennerObject("BSTISS.Rotinas")
  obj.Atualizar(CurrentSystem, CurrentQuery.FieldByName("HANDLE").AsInteger)

  Dim sql As Object
  Set sql=NewQuery
  sql.Add("SELECT SITUACAO FROM TIS_MENSAGEMTISS WHERE HANDLE=:HANDLE")
  sql.ParamByName("HANDLE").AsInteger = CurrentQuery.FieldByName("HANDLE").AsInteger
  sql.Active=True
  If sql.FieldByName("SITUACAO").AsString <> "E" Then
     obj.ProcessarRotina(CurrentSystem,CurrentQuery.FieldByName("HANDLE").AsInteger) 'Irá processar o ARQUIVO de ARQUIVO com o tipo de transação.
  End If
  Set sql= Nothing
  Set obj=Nothing
End Sub
