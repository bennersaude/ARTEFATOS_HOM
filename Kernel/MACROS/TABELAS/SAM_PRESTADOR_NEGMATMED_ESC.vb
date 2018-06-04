'HASH: 2ACB030B67F11FC8FF4D08020848D1B3
'#Uses "*bsShowMessage"
Option Explicit

Public Sub TABLE_AfterInsert()
	Dim query As BPesquisa
	Set query = NewQuery
	query.Clear
	query.Add(" SELECT MAX(VALORMAXIMO) MAXIMO FROM SAM_PRESTADOR_NEGMATMED_ESC WHERE NEGOCIACAO = :NEGOCIACAO ")
	query.ParamByName("NEGOCIACAO").AsInteger = CurrentQuery.FieldByName("NEGOCIACAO").AsInteger
	query.Active = True


	If (Not query.FieldByName("MAXIMO").IsNull) Then
		VALORMINIMO.ReadOnly = True
		CurrentQuery.FieldByName("VALORMINIMO").AsFloat = (query.FieldByName("MAXIMO").AsFloat) + 0.01
	Else
		VALORMINIMO.ReadOnly = False
	End If

	Set query = Nothing
End Sub

Public Sub TABLE_BeforeDelete(CanContinue As Boolean)
	Dim query As BPesquisa
	Set query = NewQuery
	query.Clear
	query.Add(" SELECT MAX(VALORMAXIMO) MAXIMO FROM SAM_PRESTADOR_NEGMATMED_ESC WHERE NEGOCIACAO = :NEGOCIACAO ")
	query.ParamByName("NEGOCIACAO").AsInteger = CurrentQuery.FieldByName("NEGOCIACAO").AsInteger
	query.Active = True

	If (query.FieldByName("MAXIMO").AsFloat > CurrentQuery.FieldByName("VALORMAXIMO").AsFloat) Then
		bsShowMessage("Não é possível excluir uma faixa de preços que não seja a última.", "E")
		CanContinue = False
	End If
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	If (CurrentQuery.FieldByName("PERCENTUAL").AsFloat < 0 Or CurrentQuery.FieldByName("VALORMINIMO").AsFloat < 0 Or CurrentQuery.FieldByName("VALORMAXIMO").AsFloat < 0) Then
		bsShowMessage("Valores negativos não são válidos.","E")
		CanContinue = False
		Exit Sub
	End If
	If (CurrentQuery.FieldByName("VALORMINIMO").AsFloat >= CurrentQuery.FieldByName("VALORMAXIMO").AsFloat) Then
		bsShowMessage("Valor Máximo deve ser maior que Valor Mínimo informado.", "E")
		CanContinue = False
	End If
	If (Not CurrentQuery.FieldByName("PERCENTUAL").IsNull And (CurrentQuery.FieldByName("VALORMINIMO").IsNull Or CurrentQuery.FieldByName("VALORMAXIMO").IsNull)) Then
		bsShowMessage("Informe os campos 'Valor Mínimo' e 'Valor Máximo' antes de preencher o valor do campo 'Taxa'.", "E")
		CanContinue = False
	End If
End Sub
