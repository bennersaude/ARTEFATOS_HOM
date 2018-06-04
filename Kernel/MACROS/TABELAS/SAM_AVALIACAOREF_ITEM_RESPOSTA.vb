'HASH: 8BF9B06211FD48CE93F7BB949DABB98E
'#Uses "*bsShowMessage"

Option Explicit

Dim vPontos As Integer

Public Sub TABLE_AfterEdit()
	vPontos = CurrentQuery.FieldByName("PONTOS").AsInteger
End Sub

Public Sub TABLE_BeforePost(CanContinue As Boolean)
	Dim q1 As Object
	Dim q2 As Object
	Dim vTotal As Integer
	Set q2 = NewQuery

	q2.Add("SELECT PONTUACAOMAXIMA, TIPO                 ")
	q2.Add("  FROM SAM_AVALIACAOREF_ITEM                 ")
	q2.Add(" WHERE HANDLE = :AVALIACAOREFITEM            ")

	q2.ParamByName("AVALIACAOREFITEM").Value = CurrentQuery.FieldByName("AVALIACAOREFITEM").Value
	q2.Active = True

	If CurrentQuery.FieldByName("PONTOS").AsInteger > q2.FieldByName("PONTUACAOMAXIMA").AsInteger Then
		bsShowMessage("Os pontos desta resposta ultrapassa a pontuação máxima do item!", "I")
		Exit Sub
	End If

	If CurrentQuery.FieldByName("DESCRITIVA").AsInteger = 2 Then
		CurrentQuery.FieldByName("DESCRITIVAOBRIGATORIA").AsString = "N"
	End If
End Sub
